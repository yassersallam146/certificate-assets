/* Eagle Generator – web app (phones first) */
import { getBackend } from "./backend.js?v=3";
import * as D from "./data.js?v=3";
import * as E from "./engine.js?v=3";
import { APP_VERSION } from "./config.js?v=3";

const $ = s => document.querySelector(s);
const view = $("#view");
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const today = () => new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);

let B, me = null, prof = null, L = null, lastFiles = [];

/* ── small UI helpers ─────────────────────────────────────────────────── */
function toast(msg, err = false) {
  const t = $("#toast"); t.textContent = msg; t.className = err ? "err" : ""; t.hidden = false;
  clearTimeout(toast._t); toast._t = setTimeout(() => (t.hidden = true), err ? 5000 : 2600);
}
function busy(on, text = "Working…") { $("#busy").hidden = !on; $("#busy-text").textContent = text; }
function header(title, back) {
  $("#title").textContent = title;
  $("#back").hidden = !back; $("#back").onclick = back || null;
  $("#me").hidden = !me; $("#me").onclick = () => go("profile");
}
function bar(html) {
  document.querySelector(".bar")?.remove();
  if (!html) return;
  const b = document.createElement("div"); b.className = "bar"; b.innerHTML = `<div class="inner">${html}</div>`;
  document.body.appendChild(b); return b;
}
function chips(el, values, current, onPick) {
  el.innerHTML = values.map(v => `<button type="button" class="chip ${String(v) === String(current) ? "on" : ""}" data-v="${esc(v)}">${esc(v)}</button>`).join("");
  el.onclick = e => { const b = e.target.closest(".chip"); if (!b) return;
    el.querySelectorAll(".chip").forEach(c => c.classList.toggle("on", c === b)); onPick(b.dataset.v); };
}
const steps = (n, total) => `<div class="steps">${Array.from({ length: total }, (_, i) => `<i class="${i < n ? "on" : ""}"></i>`).join("")}</div>`;
const label = e => e.name || e.id;

async function share(files, title) {
  try {
    if (navigator.canShare && navigator.canShare({ files })) { await navigator.share({ files, title }); return; }
  } catch (e) { if (e.name === "AbortError") return; }
  files.forEach(download);
}
function download(file) {
  const a = document.createElement("a"); a.href = URL.createObjectURL(file); a.download = file.name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
const DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/* ── routing ──────────────────────────────────────────────────────────── */
const screens = {};
function go(name, ...args) { bar(null); window.scrollTo(0, 0); screens[name](...args); history.replaceState(null, "", "#" + name); }

/* ═════════════════════════ LOGIN ═════════════════════════ */
screens.login = () => {
  header("Eagle Generator");
  view.innerHTML = `<div class="login">
    <img src="${D.logoUrl}" alt="" onerror="this.remove()">
    <div class="card">
      <label for="em">E-mail</label><input id="em" type="email" autocomplete="username" inputmode="email" placeholder="name@company.com">
      <label for="pw">Password</label><input id="pw" type="password" autocomplete="current-password">
      <div style="height:14px"></div>
      <button class="btn" id="go" style="width:100%">Sign in</button>
      <p class="small center" style="padding:12px 0 0"><a href="#" id="forgot">Forgot password?</a></p>
    </div>
    <p class="small muted center" style="padding:6px 0">No account? Ask the admin to create one.<br>Eagle Generator ${APP_VERSION}${B.demo ? " · DEMO" : ""}</p>
  </div>`;
  const doLogin = async () => {
    const em = $("#em").value.trim(), pw = $("#pw").value;
    if (!em || !pw) return toast("Enter your e-mail and password.", true);
    busy(true, "Signing in…");
    try { await B.signIn(em, pw); } catch (e) { busy(false); toast(e.message, true); }
  };
  $("#go").onclick = doLogin; $("#pw").onkeydown = e => { if (e.key === "Enter") doLogin(); };
  $("#forgot").onclick = async e => {
    e.preventDefault(); const em = $("#em").value.trim();
    if (!em) return toast("Type your e-mail first, then press Forgot password.", true);
    try { await B.resetPassword(em); toast("A reset link was sent to " + em); } catch (er) { toast(er.message, true); }
  };
};

/* ═════════════════════════ HOME ═════════════════════════ */
screens.home = () => {
  header("Eagle Generator");
  const admin = prof.role === "admin";
  view.innerHTML = `
    <div class="hello">Hello, <b>${esc(prof.name || me.email)}</b>${prof.empCode ? ` <span class="pill">${esc(prof.empCode)}</span>` : ""}</div>
    ${!prof.empCode || !prof.phone ? `<div class="note">Please complete your profile (Emp Code, phone) – they are printed on quotations. <a href="#" id="fix">Open profile</a></div>` : ""}
    <div class="tiles" style="margin-top:12px">
      <button class="tile t1" data-go="quote"><span class="ico">📄</span><b>Quotation</b><span>One device</span></button>
      <button class="tile t2" data-go="pkg"><span class="ico">📦</span><b>Package</b><span>Several devices</span></button>
      <button class="tile t3" data-go="cert"><span class="ico">🎓</span><b>Certificates</b><span>Training</span></button>
      <button class="tile t4" data-go="history"><span class="ico">🕘</span><b>History</b><span>${admin ? "All employees" : "My documents"}</span></button>
    </div>
    ${admin ? `<h2>Admin</h2>
      <button class="list-btn" data-go="users"><span class="ico">👥</span><span><b>Users</b><small>Add employees, block, reset passwords</small></span></button>` : ""}
    <p class="small muted center" style="padding:18px 0">Forms are loaded from your GitHub – new devices appear automatically.</p>`;
  view.querySelectorAll("[data-go]").forEach(b => (b.onclick = () => go(b.dataset.go)));
  $("#fix") && ($("#fix").onclick = e => { e.preventDefault(); go("profile"); });
};

/* ═════════════════════════ QUOTATION ═════════════════════════ */
function deviceGrid(list, selected, onPick) {
  const wrap = document.createElement("div");
  wrap.innerHTML = `<input class="search" placeholder="Search device…" type="search"><div class="devices"></div>`;
  const grid = wrap.querySelector(".devices"), search = wrap.querySelector("input");
  const draw = () => {
    const q = search.value.toLowerCase();
    grid.innerHTML = list.filter(e => label(e).toLowerCase().includes(q))
      .map(e => `<button class="dev ${e.id === selected ? "on" : ""}" data-id="${esc(e.id)}">${esc(label(e))}</button>`).join("")
      || `<div class="muted small">No device found.</div>`;
  };
  grid.onclick = ev => { const b = ev.target.closest(".dev"); if (!b) return; selected = b.dataset.id; draw(); onPick(selected); };
  search.oninput = draw; draw(); return wrap;
}

function termsFields(st) {
  return `
  <label>Client / clinic name</label><input id="acc" value="${esc(st.account)}" placeholder="e.g. مجمع العالم الطبي">
  <div class="row"><div><label>City</label><input id="city" value="${esc(st.city)}"></div>
    <div><label>Date</label><input id="date" type="date" value="${esc(st.date)}"></div></div>
  <label>Down payment %</label><div class="chips" id="dp"></div>
  <label>Installments (months)</label><div class="chips" id="ins"></div>
  <label>Notes (optional)</label><textarea id="notes" rows="3" style="min-height:80px">${esc(st.notes)}</textarea>`;
}
function bindTerms(st) {
  chips($("#dp"), ["0", "10", "20", "30", "40", "50"], st.dp, v => (st.dp = v));
  chips($("#ins"), ["0", "6", "12", "15", "18", "24"], st.ins, v => (st.ins = v));
  const keep = () => { st.account = $("#acc").value.trim(); st.city = $("#city").value.trim(); st.date = $("#date").value; st.notes = $("#notes").value.trim(); };
  ["#acc", "#city", "#date", "#notes"].forEach(s => $(s).addEventListener("input", keep));
  return keep;
}

screens.quote = (st = { step: 1, device: null, account: "", city: "", date: today(), price: "", qty: "1", dp: "20", ins: "12", notes: "" }) => {
  const back = () => (st.step > 1 ? (st.step--, screens.quote(st)) : go("home"));
  header("New quotation", back);
  if (st.step === 1) {
    view.innerHTML = `${steps(1, 3)}<h2>Choose the device</h2>`;
    view.appendChild(deviceGrid(L.quot, st.device, id => { st.device = id; $("#next").disabled = false; }));
    bar(`<button class="btn" id="next" ${st.device ? "" : "disabled"}>Next</button>`).querySelector("#next").onclick = () => { st.step = 2; screens.quote(st); };
  } else if (st.step === 2) {
    const dev = L.quot.find(e => e.id === st.device);
    view.innerHTML = `${steps(2, 3)}<h2>${esc(label(dev))}</h2>
      <div class="row"><div><label>Unit price (SAR)</label><input id="price" inputmode="decimal" value="${esc(st.price)}" placeholder="250,000"></div>
      <div><label>Qty</label><input id="qty" inputmode="numeric" value="${esc(st.qty)}"></div></div>
      ${termsFields(st)}`;
    const keep = bindTerms(st);
    bar(`<button class="btn" id="next">Review</button>`).querySelector("#next").onclick = () => {
      keep(); st.price = $("#price").value.trim(); st.qty = $("#qty").value.trim();
      try { if (E.parseAmount(st.price) <= 0) throw new Error("Enter the price."); E.parseQty(st.qty); }
      catch (e) { return toast(e.message, true); }
      if (!st.account) return toast("Enter the client name.", true);
      st.step = 3; screens.quote(st);
    };
  } else {
    const dev = L.quot.find(e => e.id === st.device);
    const price = E.parseAmount(st.price), qty = E.parseQty(st.qty), total = price * qty;
    view.innerHTML = `${steps(3, 3)}<h2>Check and create</h2>
      <div class="card"><dl class="sum">
        <dt>Device</dt><dd>${esc(label(dev))}</dd><dt>Client</dt><dd>${esc(st.account)}</dd><dt>City</dt><dd>${esc(st.city)}</dd>
        <dt>Date</dt><dd>${esc(E.longDate(st.date))}</dd><dt>Price × Qty</dt><dd>${E.fmt(price)} × ${qty}</dd>
        <dt>Down payment</dt><dd>${esc(st.dp)}%</dd><dt>Installments</dt><dd>${esc(st.ins)} months</dd></dl>
        <div class="total">${E.fmt(total)} SAR</div><div class="small muted" style="text-align:right">${esc(E.wordsSAR(total))}</div></div>
      ${price < 5000 ? `<div class="note">⚠ The price looks very low (${E.fmt(price)} SAR). Check it before creating.</div>` : ""}
      <label>Quotation number (serial)</label><input id="serial" inputmode="numeric" value="${esc(prof.quotNext || "10001")}">
      <p class="small muted">Printed as ${esc(prof.empCode || "CODE")}#number#. You can type a different start number – the next one continues from it.</p>`;
    bar(`<button class="btn green" id="make">Create quotation</button>`).querySelector("#make").onclick = async () => {
      const serial = $("#serial").value.trim();
      if (!/^\d+$/.test(serial)) return toast("The serial must be a number.", true);
      busy(true, "Creating quotation…");
      try {
        const f = { kind: "quotation", device: dev.id, deviceLabel: label(dev), account: st.account, city: st.city, date: st.date,
          price: st.price, qty: st.qty, dp: st.dp, installments: st.ins, notes: st.notes, serial };
        const file = await buildQuotation(f, dev);
        await B.takeSerials(me.uid, "quotNext", serial, 1);
        prof.quotNext = String(+serial + 1).padStart(serial.length, "0");
        await B.addRegister({ uid: me.uid, by: prof.name || me.email, empCode: prof.empCode || "", ...f, total, file: file.name });
        busy(false); go("done", [file], "Quotation created");
      } catch (e) { busy(false); toast(e.message, true); }
    };
  }
};

async function buildQuotation(f, dev) {
  const zip = await E.openDocx(await D.quotationForm(dev));
  const price = E.parseAmount(f.price), qty = E.parseQty(f.qty), dp = E.parseAmount(f.dp);
  await E.fill(zip, E.quotationMap({ ...f, empCode: prof.empCode || "", empName: prof.name || "", empPhone: prof.phone || "",
    installments: f.installments }, price * qty, price, qty, dp));
  const name = `${E.safeName(f.deviceLabel)}_Quotation_${E.safeName(f.account)}_${f.serial}.docx`;
  return new File([await E.toBlob(zip)], name, { type: DOCX });
}

/* ═════════════════════════ PACKAGE ═════════════════════════ */
screens.pkg = (st = { step: 1, sel: {}, account: "", city: "", date: today(), dp: "20", ins: "12", notes: "" }) => {
  const back = () => (st.step > 1 ? (st.step--, screens.pkg(st)) : go("home"));
  header("Package quotation", back);
  const devs = L.pkg.length ? L.pkg : L.quot.map(e => ({ id: e.id, name: e.name, desc: "" }));
  if (st.step === 1) {
    view.innerHTML = `${steps(1, 3)}<h2>Tick devices, enter price and qty</h2><div id="items"></div>`;
    $("#items").innerHTML = devs.map(d => {
      const s = st.sel[d.id];
      return `<div class="pkg-item" data-id="${esc(d.id)}"><input type="checkbox" ${s ? "checked" : ""}>
        <div><b>${esc(label(d))}</b><div class="small muted">${esc((d.desc || "").slice(0, 90))}</div></div>
        <div class="fields" ${s ? "" : "hidden"}><input class="p" inputmode="decimal" placeholder="Price SAR" value="${esc(s?.price || "")}">
        <input class="q" inputmode="numeric" placeholder="Qty" value="${esc(s?.qty || "1")}"></div></div>`;
    }).join("");
    $("#items").onclick = e => {
      const it = e.target.closest(".pkg-item"); if (!it || e.target.tagName === "INPUT" && e.target.type !== "checkbox") return;
      const cb = it.querySelector("[type=checkbox]"); if (e.target !== cb) cb.checked = !cb.checked;
      it.querySelector(".fields").hidden = !cb.checked;
      if (cb.checked) st.sel[it.dataset.id] = st.sel[it.dataset.id] || { price: "", qty: "1" }; else delete st.sel[it.dataset.id];
    };
    $("#items").oninput = e => { const it = e.target.closest(".pkg-item"); const s = st.sel[it.dataset.id]; if (!s) return;
      s.price = it.querySelector(".p").value.trim(); s.qty = it.querySelector(".q").value.trim(); };
    bar(`<button class="btn" id="next">Next</button>`).querySelector("#next").onclick = () => {
      const ids = Object.keys(st.sel);
      if (!ids.length) return toast("Tick at least one device.", true);
      try { ids.forEach(id => { if (E.parseAmount(st.sel[id].price) <= 0) throw new Error(`Enter the price for ${label(devs.find(d => d.id === id))}.`); E.parseQty(st.sel[id].qty); }); }
      catch (e) { return toast(e.message, true); }
      st.step = 2; screens.pkg(st);
    };
  } else if (st.step === 2) {
    view.innerHTML = `${steps(2, 3)}<h2>Client and payment</h2>${termsFields(st)}`;
    const keep = bindTerms(st);
    bar(`<button class="btn" id="next">Review</button>`).querySelector("#next").onclick = () => {
      keep(); if (!st.account) return toast("Enter the client name.", true); st.step = 3; screens.pkg(st); };
  } else {
    const rows = Object.entries(st.sel).map(([id, s]) => { const d = devs.find(x => x.id === id);
      return { id, label: label(d), desc: d.desc || "", price: E.parseAmount(s.price), qty: E.parseQty(s.qty) }; });
    const gt = rows.reduce((a, r) => a + r.price * r.qty, 0);
    view.innerHTML = `${steps(3, 3)}<h2>Check and create</h2>
      <div class="card"><dl class="sum">${rows.map(r => `<dt>${esc(r.label)}</dt><dd>${E.fmt(r.price)} × ${r.qty}</dd>`).join("")}
        <dt>Client</dt><dd>${esc(st.account)}</dd><dt>Down payment</dt><dd>${esc(st.dp)}%</dd><dt>Installments</dt><dd>${esc(st.ins)} months</dd></dl>
        <div class="total">${E.fmt(gt)} SAR</div></div>
      ${rows.some(r => r.price < 5000) ? `<div class="note">⚠ Some prices look very low. Check them before creating.</div>` : ""}
      <label>Quotation number (serial)</label><input id="serial" inputmode="numeric" value="${esc(prof.quotNext || "10001")}">`;
    bar(`<button class="btn green" id="make">Create package quotation</button>`).querySelector("#make").onclick = async () => {
      const serial = $("#serial").value.trim();
      if (!/^\d+$/.test(serial)) return toast("The serial must be a number.", true);
      busy(true, "Creating package quotation…");
      try {
        const f = { kind: "package", items: rows, account: st.account, city: st.city, date: st.date, dp: st.dp,
          installments: st.ins, notes: st.notes, serial };
        const file = await buildPackage(f);
        await B.takeSerials(me.uid, "quotNext", serial, 1);
        prof.quotNext = String(+serial + 1).padStart(serial.length, "0");
        await B.addRegister({ uid: me.uid, by: prof.name || me.email, empCode: prof.empCode || "", ...f,
          deviceLabel: rows.map(r => r.label).join(" + "), total: gt, file: file.name });
        busy(false); go("done", [file], "Package quotation created");
      } catch (e) { busy(false); toast(e.message, true); }
    };
  }
};

async function buildPackage(f) {
  const zip = await E.openDocx(await D.packageForm());
  const gt = f.items.reduce((a, r) => a + r.price * r.qty, 0);
  await E.fill(zip, E.quotationMap({ ...f, empCode: prof.empCode || "", empName: prof.name || "", empPhone: prof.phone || "" },
    gt, 0, 0, E.parseAmount(f.dp)));
  await E.packageRows(zip, f.items, gt);
  return new File([await E.toBlob(zip)], `Package_Quotation_${E.safeName(f.account)}_${f.serial}.docx`, { type: DOCX });
}

/* ═════════════════════════ CERTIFICATES ═════════════════════════ */
screens.cert = (st = { step: 1, framed: true, device: null, from: today(), to: today(), city: "", names: "", stamp: true, qr: true }) => {
  const back = () => (st.step > 1 ? (st.step--, screens.cert(st)) : go("home"));
  header("Training certificates", back);
  const list = st.framed ? L.framed : L.unframed;
  if (st.step === 1) {
    view.innerHTML = `${steps(1, 3)}<h2>Certificate style</h2><div class="chips" id="frame"></div><h2>Device</h2>`;
    chips($("#frame"), ["With frame", "No frame"], st.framed ? "With frame" : "No frame", v => { st.framed = v === "With frame"; st.device = null; screens.cert(st); });
    view.appendChild(deviceGrid(list, st.device, id => { st.device = id; $("#next").disabled = false; }));
    bar(`<button class="btn" id="next" ${st.device ? "" : "disabled"}>Next</button>`).querySelector("#next").onclick = () => { st.step = 2; screens.cert(st); };
  } else if (st.step === 2) {
    view.innerHTML = `${steps(2, 3)}<h2>Training details</h2>
      <div class="row"><div><label>From</label><input id="from" type="date" value="${esc(st.from)}"></div><div><label>To</label><input id="to" type="date" value="${esc(st.to)}"></div></div>
      <label>City</label><input id="city" value="${esc(st.city)}">
      <label>Trainee names – one per line</label><textarea id="names" placeholder="Paste names here (from WhatsApp, Excel …)">${esc(st.names)}</textarea>
      <div style="margin-top:8px"><button class="btn sec btn-inline" id="xl">📥 Import from Excel</button><input type="file" id="xlf" accept=".xlsx,.xls,.csv" hidden></div>
      <label class="toggle"><input type="checkbox" id="stamp" ${st.stamp ? "checked" : ""}> Add the company stamp</label>
      <label class="toggle"><input type="checkbox" id="qr" ${st.qr ? "checked" : ""}> Add QR code with the certificate details</label>`;
    $("#xl").onclick = () => $("#xlf").click();
    $("#xlf").onchange = async e => {
      const file = e.target.files[0]; if (!file) return;
      try {
        const wb = window.XLSX.read(await file.arrayBuffer());
        const rows = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
        let names = rows.map(r => String(r[0] ?? "").trim()).filter(Boolean);
        if (names.length && /^(name|names|full name|trainee|trainees|الاسم|اسم|الأسماء|الاسماء|اسم المتدرب)$/i.test(names[0])) names = names.slice(1);
        $("#names").value = names.join("\n"); toast(`${names.length} names imported`);
      } catch (er) { toast("Could not read the file: " + er.message, true); }
    };
    bar(`<button class="btn" id="next">Review</button>`).querySelector("#next").onclick = () => {
      st.from = $("#from").value; st.to = $("#to").value; st.city = $("#city").value.trim();
      st.names = $("#names").value; st.stamp = $("#stamp").checked; st.qr = $("#qr").checked;
      const names = st.names.split(/[\n,]+/).map(s => s.trim()).filter(Boolean);
      if (!names.length) return toast("Add at least one trainee name.", true);
      if (!st.city) return toast("Enter the city.", true);
      st.step = 3; screens.cert(st);
    };
  } else {
    const dev = list.find(e => e.id === st.device);
    const names = st.names.split(/[\n,]+/).map(s => s.trim()).filter(Boolean);
    view.innerHTML = `${steps(3, 3)}<h2>Check and create</h2>
      <div class="card"><dl class="sum"><dt>Device</dt><dd>${esc(label(dev))}</dd><dt>Style</dt><dd>${st.framed ? "With frame" : "No frame"}</dd>
      <dt>Dates</dt><dd>${E.shortDate(st.from)} – ${E.shortDate(st.to)}</dd><dt>City</dt><dd>${esc(st.city)}</dd>
      <dt>Trainees</dt><dd>${names.length}</dd><dt>Stamp / QR</dt><dd>${st.stamp ? "✓" : "–"} / ${st.qr ? "✓" : "–"}</dd></dl></div>
      <label>First certificate number</label><input id="serial" inputmode="numeric" value="${esc(prof.certNext || "0001")}">
      <p class="small muted">Printed as ${esc(prof.empCode || "CODE")}-number (e.g. ${esc(prof.empCode || "YS")}-${esc(prof.certNext || "0001")}), then +1 for each trainee.</p>`;
    bar(`<button class="btn green" id="make">Create ${names.length} certificate${names.length > 1 ? "s" : ""}</button>`).querySelector("#make").onclick = async () => {
      const serial = $("#serial").value.trim();
      if (!/^\d+$/.test(serial)) return toast("The number must be digits only (e.g. 0042).", true);
      busy(true, "Creating certificates…");
      try {
        await B.takeSerials(me.uid, "certNext", serial, names.length);   // reserve all numbers first
        prof.certNext = String(+serial + names.length).padStart(serial.length, "0");
        const form = await D.certificateForm(dev, st.framed);
        const stampBytes = st.stamp ? await D.stampImage() : null;
        const files = [], failed = [];
        for (let i = 0; i < names.length; i++) {
          $("#busy-text").textContent = `Creating ${i + 1} / ${names.length}…`;
          const no = String(+serial + i).padStart(serial.length, "0");
          const certNo = prof.empCode ? `${prof.empCode}-${no}` : no;
          try {
            const file = await buildCertificate({ form, stampBytes, qr: st.qr, name: names[i], certNo, dev, from: st.from, to: st.to, city: st.city });
            files.push(file);
            await B.addRegister({ uid: me.uid, by: prof.name || me.email, empCode: prof.empCode || "", kind: "certificate",
              serial: certNo, device: dev.id, deviceLabel: label(dev), framed: st.framed, trainee: names[i],
              from: st.from, to: st.to, city: st.city, stamp: st.stamp, qr: st.qr, file: file.name });
          } catch (e) { failed.push(`${names[i]}: ${e.message}`); }
        }
        busy(false);
        if (failed.length) toast(`${failed.length} not created – ${failed[0]}`, true);
        go("done", files, `${files.length} certificate${files.length > 1 ? "s" : ""} created`);
      } catch (e) { busy(false); toast(e.message, true); }
    };
  }
};

async function buildCertificate({ form, stampBytes, qr, name, certNo, dev, from, to, city }) {
  const zip = await E.openDocx(form.slice(0));
  await E.fill(zip, { "{{NAME}}": name, "{{DEVICE}}": label(dev), "{{DATE}}": `from ${E.shortDate(from)} to ${E.shortDate(to)}`,
    "{{CITY}}": city, "{{SERIAL}}": certNo });
  if (stampBytes) await E.stamp(zip, stampBytes);
  if (qr) await E.addFooterQR(zip, await E.qrPng(`AME-KSA Training Certificate\nNo: ${certNo}\nName: ${name}\nDevice: ${label(dev)}\nFrom ${E.shortDate(from)} to ${E.shortDate(to)}\nCity: ${city}`));
  return new File([await E.toBlob(zip)], `${E.safeName(label(dev))}_Certificate_${E.safeName(name)}.docx`, { type: DOCX });
}

/* ═════════════════════════ DONE / SHARE ═════════════════════════ */
screens.done = (files = lastFiles, title = "Done") => {
  lastFiles = files;
  header(title, () => go("home"));
  view.innerHTML = `<div class="done"><div class="big">✅</div><h2 style="text-align:center">${esc(title)}</h2>
    <p class="small muted">Share the Word file on WhatsApp or e-mail, or save it to your phone.</p></div>
    ${files.map((f, i) => `<div class="file"><span>📄</span><span style="flex:1">${esc(f.name)}</span><button class="btn sec btn-inline" data-i="${i}">Save</button></div>`).join("")}
    ${files.length > 1 ? `<button class="btn sec" id="zip" style="width:100%;margin-top:6px">Download all as one ZIP</button>` : ""}
    <div class="note" style="margin-top:14px">To send as <b>PDF</b>: open the file in the <b>Word</b> app on your phone → Share / Export → <b>PDF</b>. (Converting in the browser would move the pictures.)</div>`;
  view.querySelectorAll("[data-i]").forEach(b => (b.onclick = () => download(files[+b.dataset.i])));
  $("#zip") && ($("#zip").onclick = async () => {
    const z = new window.JSZip(); files.forEach(f => z.file(f.name, f));
    download(new File([await z.generateAsync({ type: "blob" })], `Certificates_${today()}.zip`, { type: "application/zip" }));
  });
  bar(`<button class="btn sec" id="home">Home</button><button class="btn green" id="share">Share${files.length > 1 ? ` (${files.length})` : ""}</button>`);
  $("#home").onclick = () => go("home");
  $("#share").onclick = () => share(files, title);
};

/* ═════════════════════════ HISTORY ═════════════════════════ */
screens.history = async () => {
  header(prof.role === "admin" ? "History – all" : "My history", () => go("home"));
  view.innerHTML = `<div class="center muted">Loading…</div>`;
  let rows;
  try { rows = await B.listRegister(me.uid, prof.role === "admin"); }
  catch (e) { view.innerHTML = `<div class="note">${esc(e.message)}</div>`; return; }
  const kindPill = k => k === "certificate" ? `<span class="pill c">Certificate</span>` : k === "package" ? `<span class="pill p">Package</span>` : `<span class="pill q">Quotation</span>`;
  view.innerHTML = `<input class="search" type="search" id="q" placeholder="Search client, device, name, number…">
    <div class="small muted" id="cnt" style="margin:6px 2px"></div><div id="lst"></div>`;
  const draw = () => {
    const q = $("#q").value.toLowerCase();
    const f = rows.filter(r => JSON.stringify([r.account, r.deviceLabel, r.trainee, r.serial, r.by, r.city]).toLowerCase().includes(q));
    $("#cnt").textContent = `${f.length} document${f.length === 1 ? "" : "s"}`;
    $("#lst").innerHTML = f.map(r => `<div class="card hist" data-id="${esc(r.id)}">
      <div class="t"><span>${kindPill(r.kind)} <b>${esc(r.kind === "certificate" ? r.trainee : r.account)}</b></span><span class="small muted">${new Date(r.createdAt).toLocaleDateString("en-GB")}</span></div>
      <div class="small">${esc(r.deviceLabel)} · No. ${esc(r.serial)}${r.total ? ` · ${E.fmt(r.total)} SAR` : ""}</div>
      <div class="small muted">${prof.role === "admin" ? esc(r.by) + " · " : ""}${esc(r.city || "")}</div>
      <button class="btn sec btn-inline" data-re="${esc(r.id)}" style="margin-top:8px">Create again</button></div>`).join("") || `<div class="center muted">Nothing yet.</div>`;
  };
  $("#q").oninput = draw; draw();
  $("#lst").onclick = async e => {
    const b = e.target.closest("[data-re]"); if (!b) return;
    const r = rows.find(x => x.id === b.dataset.re);
    busy(true, "Creating again…");
    try { go("done", [await rebuild(r)], "Created again"); } catch (er) { toast(er.message, true); } finally { busy(false); }
  };
  if (prof.role === "admin") bar(`<button class="btn sec" id="csv">⬇ Export to Excel (CSV)</button>`).querySelector("#csv").onclick = () => {
    const cols = ["createdAt", "kind", "serial", "by", "empCode", "account", "city", "deviceLabel", "trainee", "qty", "price", "total", "dp", "installments", "from", "to", "file"];
    const csv = "﻿" + [cols.join(",")].concat(rows.map(r => cols.map(c => {
      let v = c === "createdAt" ? new Date(r[c]).toLocaleString("en-GB") : (r[c] ?? "");
      return `"${String(v).replace(/"/g, '""')}"`; }).join(","))).join("\n");
    download(new File([csv], `Eagle_Register_${today()}.csv`, { type: "text/csv" }));
  };
};

async function rebuild(r) {
  if (r.kind === "quotation") return buildQuotation(r, L.quot.find(e => e.id === r.device) || { id: r.device });
  if (r.kind === "package") return buildPackage(r);
  const list = r.framed ? L.framed : L.unframed;
  const dev = list.find(e => e.id === r.device) || { id: r.device, name: r.deviceLabel };
  return buildCertificate({ form: await D.certificateForm(dev, r.framed), stampBytes: r.stamp ? await D.stampImage() : null,
    qr: r.qr, name: r.trainee, certNo: r.serial, dev, from: r.from, to: r.to, city: r.city });
}

/* ═════════════════════════ PROFILE ═════════════════════════ */
screens.profile = () => {
  header("My profile", () => go("home"));
  view.innerHTML = `<div class="card">
      <label>Full name (printed on quotations)</label><input id="nm" value="${esc(prof.name)}">
      <div class="row"><div><label>Emp Code</label><input id="code" value="${esc(prof.empCode)}" placeholder="e.g. YS"></div>
      <div><label>Mobile</label><input id="ph" inputmode="tel" value="${esc(prof.phone)}"></div></div>
      <div style="height:12px"></div><button class="btn" id="save" style="width:100%">Save</button></div>
    <h2>Change password</h2><div class="card">
      <label>Current password</label><input id="o" type="password" autocomplete="current-password">
      <label>New password</label><input id="n1" type="password" autocomplete="new-password">
      <label>Repeat new password</label><input id="n2" type="password" autocomplete="new-password">
      <div style="height:12px"></div><button class="btn sec" id="pw" style="width:100%">Change password</button></div>
    <p class="small muted center">${esc(me.email)} · ${prof.role === "admin" ? "Admin" : "User"}</p>`;
  $("#save").onclick = async () => {
    const data = { name: $("#nm").value.trim(), empCode: $("#code").value.trim().toUpperCase(), phone: $("#ph").value.trim() };
    try { await B.updateProfile(me.uid, data); Object.assign(prof, data); toast("Saved ✓"); } catch (e) { toast(e.message, true); }
  };
  $("#pw").onclick = async () => {
    const o = $("#o").value, a = $("#n1").value, b = $("#n2").value;
    if (a.length < 6) return toast("Use at least 6 characters.", true);
    if (a !== b) return toast("The new passwords do not match.", true);
    try { await B.changePassword(o, a); toast("Password changed ✓"); ["#o", "#n1", "#n2"].forEach(s => ($(s).value = "")); }
    catch (e) { toast(e.message, true); }
  };
  bar(`<button class="btn red" id="out">Sign out</button>`).querySelector("#out").onclick = () => B.signOut();
};

/* ═════════════════════════ USERS (admin) ═════════════════════════ */
screens.users = async () => {
  header("Users", () => go("home"));
  view.innerHTML = `<div class="center muted">Loading…</div>`;
  const users = (await B.listUsers()).sort((a, b) => (a.name || a.email).localeCompare(b.name || b.email));
  view.innerHTML = users.map(u => `<div class="card"><div class="user">
      <div><b>${esc(u.name || u.email)}</b> ${u.role === "admin" ? `<span class="pill">Admin</span>` : ""} ${u.active === false ? `<span class="pill" style="background:#FFEBEE;color:#B71C1C">Blocked</span>` : ""}
      <div class="small muted">${esc(u.email)}${u.empCode ? " · " + esc(u.empCode) : ""}</div></div></div>
      ${u.uid !== me.uid ? `<div class="chips" style="margin-top:10px">
        <button class="chip" data-a="active" data-u="${u.uid}">${u.active === false ? "Unblock" : "Block"}</button>
        <button class="chip" data-a="role" data-u="${u.uid}">${u.role === "admin" ? "Make user" : "Make admin"}</button>
        <button class="chip" data-a="reset" data-u="${u.uid}">Send password reset</button></div>` : ""}</div>`).join("");
  view.onclick = async e => {
    const b = e.target.closest("[data-a]"); if (!b) return;
    const u = users.find(x => x.uid === b.dataset.u);
    try {
      if (b.dataset.a === "active") await B.updateProfile(u.uid, { active: u.active === false });
      if (b.dataset.a === "role") await B.updateProfile(u.uid, { role: u.role === "admin" ? "user" : "admin" });
      if (b.dataset.a === "reset") { await B.resetPassword(u.email); toast("Reset link sent to " + u.email); return; }
      toast("Saved ✓"); screens.users();
    } catch (er) { toast(er.message, true); }
  };
  bar(`<button class="btn teal" id="add">＋ Add employee</button>`).querySelector("#add").onclick = () => go("adduser");
};

screens.adduser = () => {
  header("Add employee", () => go("users"));
  view.innerHTML = `<div class="card">
    <label>Full name</label><input id="nm">
    <label>E-mail (used to sign in)</label><input id="em" type="email" inputmode="email">
    <label>First password (they can change it)</label><input id="pw" type="text" value="${Math.random().toString(36).slice(2, 6)}${Math.floor(1000 + Math.random() * 9000)}">
    <div class="row"><div><label>Emp Code</label><input id="code" placeholder="e.g. AH"></div><div><label>Mobile</label><input id="ph" inputmode="tel"></div></div>
    <label class="toggle"><input type="checkbox" id="adm"> Admin (can manage users and see everyone's history)</label></div>
    <p class="small muted">Send the e-mail and first password to the employee (e.g. on WhatsApp) with the app link.</p>`;
  bar(`<button class="btn teal" id="mk">Create account</button>`).querySelector("#mk").onclick = async () => {
    const em = $("#em").value.trim(), pw = $("#pw").value.trim(), nm = $("#nm").value.trim();
    if (!nm || !/@/.test(em)) return toast("Enter the name and a valid e-mail.", true);
    if (pw.length < 6) return toast("The password needs at least 6 characters.", true);
    busy(true, "Creating account…");
    try {
      await B.createUser(em, pw, { name: nm, empCode: $("#code").value.trim().toUpperCase(), phone: $("#ph").value.trim(), role: $("#adm").checked ? "admin" : "user" });
      busy(false);
      const msg = `Eagle Generator\nLink: ${location.origin + location.pathname}\nE-mail: ${em}\nPassword: ${pw}\n(Change it in Profile after the first login.)`;
      view.innerHTML = `<div class="done"><div class="big">✅</div><h2 style="text-align:center">Account created</h2></div>
        <div class="card"><pre style="white-space:pre-wrap;margin:0;font:inherit">${esc(msg)}</pre></div>`;
      bar(`<button class="btn sec" id="b">Users</button><button class="btn green" id="s">Send to employee</button>`);
      $("#b").onclick = () => go("users");
      $("#s").onclick = async () => { try { await navigator.share({ text: msg }); } catch { await navigator.clipboard?.writeText(msg); toast("Copied – paste it in WhatsApp"); } };
    } catch (e) { busy(false); toast(e.message, true); }
  };
};

/* ═════════════════════════ START ═════════════════════════ */
(async () => {
  try { B = await getBackend(); }
  catch (e) { view.innerHTML = `<div class="note">Could not start: ${esc(e.message)}</div>`; return; }
  B.onAuth(async u => {
    busy(false);
    me = u; prof = null;
    if (!u) return go("login");
    busy(true, "Loading…");
    try {
      prof = await B.getProfile(u.uid);
      if (!prof) { await B.signOut(); busy(false); return toast("Your account is not set up yet – ask the admin.", true); }
      if (prof.active === false) { await B.signOut(); busy(false); return toast("Your account is blocked – contact the admin.", true); }
      L = await D.lists();
      busy(false); go("home");
    } catch (e) { busy(false); toast(e.message, true); view.innerHTML = `<div class="note">${esc(e.message)}</div>`; }
  });
  if ("serviceWorker" in navigator && !B.demo) navigator.serviceWorker.register("sw.js").catch(() => {});
})();
