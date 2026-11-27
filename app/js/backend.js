/* Backend: Firebase (real) or Demo (in-browser, for testing: add ?demo=1).
 * Same functions in both, so the screens do not care which one is used.
 *
 * Firestore layout
 *   users/{uid}     { name, email, role: "admin"|"user", active, empCode, phone, quotNext, certNext }
 *   register/{id}   { uid, by, kind: "quotation"|"package"|"certificate", serial, createdAt, ...fields }
 */
import { FIREBASE, OWNER_EMAIL } from "./config.js?v=4";

const DEMO = new URLSearchParams(location.search).has("demo");

/* ════════════════════════ DEMO ════════════════════════ */
function demoBackend() {
  const load = () => JSON.parse(localStorage.getItem("eagle_demo") || '{"users":{},"register":[],"pw":{}}');
  const save = db => localStorage.setItem("eagle_demo", JSON.stringify(db));
  let current = null, listeners = [];
  const db0 = load();
  if (!Object.keys(db0.users).length) {
    db0.users["u_admin"] = { name: "Dr. Yasser Sallam", email: "admin@demo", role: "admin", active: true,
      empCode: "YS", phone: "0540211533", quotNext: "10100", certNext: "0042" };
    db0.pw["admin@demo"] = { uid: "u_admin", pw: "demo1234" };
    save(db0);
  }
  return {
    demo: true,
    onAuth(cb) { listeners.push(cb); setTimeout(() => cb(current), 0); },
    async signIn(email, pw) {
      const db = load(), a = db.pw[email.toLowerCase()];
      if (!a || a.pw !== pw) throw new Error("Wrong e-mail or password.");
      current = { uid: a.uid, email }; listeners.forEach(f => f(current));
    },
    async signOut() { current = null; listeners.forEach(f => f(null)); },
    async resetPassword() { },
    async changePassword(oldPw, newPw) {
      const db = load(), a = db.pw[current.email];
      if (a.pw !== oldPw) throw new Error("Current password is not correct.");
      a.pw = newPw; save(db);
    },
    async getProfile(uid) { return load().users[uid] || null; },
    async updateProfile(uid, data) { const db = load(); Object.assign(db.users[uid], data); save(db); },
    async listUsers() { const u = load().users; return Object.entries(u).map(([uid, v]) => ({ uid, ...v })); },
    async createUser(email, pw, data) {
      const db = load(); const uid = "u_" + Math.random().toString(36).slice(2, 8);
      db.users[uid] = { ...data, email: email.toLowerCase(), active: true, quotNext: "10001", certNext: "0001" };
      db.pw[email.toLowerCase()] = { uid, pw }; save(db); return uid;
    },
    async takeSerials(uid, field, start, count) {
      const db = load(); const used = start;
      db.users[uid][field] = String(+start + count).padStart(String(start).length, "0"); save(db); return used;
    },
    async addRegister(entry) { const db = load(); db.register.unshift({ id: "r" + Date.now(), ...entry, createdAt: Date.now() }); save(db); },
    async listRegister(uid, isAdmin) { const r = load().register; return isAdmin ? r : r.filter(x => x.uid === uid); },
  };
}

/* ════════════════════════ FIREBASE ════════════════════════ */
async function firebaseBackend() {
  const { initializeApp, deleteApp } = await import("../lib/firebase-app.js");
  const A = await import("../lib/firebase-auth.js");
  const F = await import("../lib/firebase-firestore.js");
  const app = initializeApp(FIREBASE);
  const auth = A.getAuth(app);
  await A.setPersistence(auth, A.browserLocalPersistence);
  const db = F.getFirestore(app);
  return {
    demo: false,
    onAuth(cb) { A.onAuthStateChanged(auth, u => cb(u ? { uid: u.uid, email: u.email } : null)); },
    async signIn(email, pw) {
      try { await A.signInWithEmailAndPassword(auth, email.trim(), pw); }
      catch (e) { throw new Error(/invalid|wrong|not-found|credential/i.test(e.code) ? "Wrong e-mail or password." : e.message); }
    },
    async signOut() { await A.signOut(auth); },
    async resetPassword(email) { await A.sendPasswordResetEmail(auth, email.trim()); },
    async changePassword(oldPw, newPw) {
      const u = auth.currentUser;
      try { await A.reauthenticateWithCredential(u, A.EmailAuthProvider.credential(u.email, oldPw)); }
      catch { throw new Error("Current password is not correct."); }
      await A.updatePassword(u, newPw);
    },
    async getProfile(uid) {
      const s = await F.getDoc(F.doc(db, "users", uid));
      if (s.exists()) return s.data();
      const u = auth.currentUser;
      if (u && u.email && u.email.toLowerCase() === OWNER_EMAIL.toLowerCase()) {      // owner, first start
        const prof = { name: "Admin", email: u.email.toLowerCase(), role: "admin", active: true,
          empCode: "", phone: "", quotNext: "10001", certNext: "0001" };
        await F.setDoc(F.doc(db, "users", uid), prof);
        return prof;
      }
      return null;
    },
    async updateProfile(uid, data) { await F.updateDoc(F.doc(db, "users", uid), data); },
    async listUsers() {
      const q = await F.getDocs(F.collection(db, "users"));
      return q.docs.map(d => ({ uid: d.id, ...d.data() }));
    },
    async createUser(email, pw, data) {
      // a second Firebase instance, so the admin stays signed in
      const tmp = initializeApp(FIREBASE, "create-" + Date.now());
      try {
        const cred = await A.createUserWithEmailAndPassword(A.getAuth(tmp), email.trim(), pw);
        await F.setDoc(F.doc(db, "users", cred.user.uid), { ...data, email: email.trim().toLowerCase(),
          active: true, quotNext: "10001", certNext: "0001" });
        await A.signOut(A.getAuth(tmp));
        return cred.user.uid;
      } catch (e) {
        throw new Error(/email-already-in-use/.test(e.code) ? "This e-mail already has an account." :
          /weak-password/.test(e.code) ? "Password must be at least 6 characters." : e.message);
      } finally { await deleteApp(tmp); }
    },
    /* Reserve `count` numbers starting at `start` (the number the user typed)
       and store the next free one – one transaction, so nothing is given twice. */
    async takeSerials(uid, field, start, count) {
      const ref = F.doc(db, "users", uid);
      return await F.runTransaction(db, async tr => {
        const width = String(start).length;
        tr.update(ref, { [field]: String(+start + count).padStart(width, "0") });
        return start;
      });
    },
    async addRegister(entry) {
      await F.addDoc(F.collection(db, "register"), { ...entry, createdAt: F.serverTimestamp() });
    },
    async listRegister(uid, isAdmin) {
      const col = F.collection(db, "register");
      const q = isAdmin ? F.query(col, F.orderBy("createdAt", "desc"), F.limit(300))
                        : F.query(col, F.where("uid", "==", uid), F.limit(300));
      const s = await F.getDocs(q);
      const rows = s.docs.map(d => ({ id: d.id, ...d.data(),
        createdAt: d.data().createdAt?.toMillis ? d.data().createdAt.toMillis() : Date.now() }));
      return rows.sort((a, b) => b.createdAt - a.createdAt);
    },
  };
}

export async function getBackend() {
  return DEMO ? demoBackend() : await firebaseBackend();
}
