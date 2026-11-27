/* Arabic letter joining (initial / medial / final / isolated forms + lam-alef),
 * so Arabic names print correctly in PDFs on any phone. */
// letter: [isolated, final, initial, medial]   (null = letter does not join to the next one)
const F = {
  "ء": [0xFE80, 0xFE80, null, null], "آ": [0xFE81, 0xFE82, null, null], "أ": [0xFE83, 0xFE84, null, null],
  "ؤ": [0xFE85, 0xFE86, null, null], "إ": [0xFE87, 0xFE88, null, null], "ئ": [0xFE89, 0xFE8A, 0xFE8B, 0xFE8C],
  "ا": [0xFE8D, 0xFE8E, null, null], "ب": [0xFE8F, 0xFE90, 0xFE91, 0xFE92], "ة": [0xFE93, 0xFE94, null, null],
  "ت": [0xFE95, 0xFE96, 0xFE97, 0xFE98], "ث": [0xFE99, 0xFE9A, 0xFE9B, 0xFE9C], "ج": [0xFE9D, 0xFE9E, 0xFE9F, 0xFEA0],
  "ح": [0xFEA1, 0xFEA2, 0xFEA3, 0xFEA4], "خ": [0xFEA5, 0xFEA6, 0xFEA7, 0xFEA8], "د": [0xFEA9, 0xFEAA, null, null],
  "ذ": [0xFEAB, 0xFEAC, null, null], "ر": [0xFEAD, 0xFEAE, null, null], "ز": [0xFEAF, 0xFEB0, null, null],
  "س": [0xFEB1, 0xFEB2, 0xFEB3, 0xFEB4], "ش": [0xFEB5, 0xFEB6, 0xFEB7, 0xFEB8], "ص": [0xFEB9, 0xFEBA, 0xFEBB, 0xFEBC],
  "ض": [0xFEBD, 0xFEBE, 0xFEBF, 0xFEC0], "ط": [0xFEC1, 0xFEC2, 0xFEC3, 0xFEC4], "ظ": [0xFEC5, 0xFEC6, 0xFEC7, 0xFEC8],
  "ع": [0xFEC9, 0xFECA, 0xFECB, 0xFECC], "غ": [0xFECD, 0xFECE, 0xFECF, 0xFED0], "ف": [0xFED1, 0xFED2, 0xFED3, 0xFED4],
  "ق": [0xFED5, 0xFED6, 0xFED7, 0xFED8], "ك": [0xFED9, 0xFEDA, 0xFEDB, 0xFEDC], "ل": [0xFEDD, 0xFEDE, 0xFEDF, 0xFEE0],
  "م": [0xFEE1, 0xFEE2, 0xFEE3, 0xFEE4], "ن": [0xFEE5, 0xFEE6, 0xFEE7, 0xFEE8], "ه": [0xFEE9, 0xFEEA, 0xFEEB, 0xFEEC],
  "و": [0xFEED, 0xFEEE, null, null], "ى": [0xFEEF, 0xFEF0, null, null], "ي": [0xFEF1, 0xFEF2, 0xFEF3, 0xFEF4],
  "ـ": [0x0640, 0x0640, 0x0640, 0x0640],
};
const LAM_ALEF = { "آ": [0xFEF5, 0xFEF6], "أ": [0xFEF7, 0xFEF8], "إ": [0xFEF9, 0xFEFA], "ا": [0xFEFB, 0xFEFC] };
const HARAKAT = /[ً-ٰٟۖ-ۭ]/g;
const joinsNext = c => F[c] && F[c][2] !== null;

/** Arabic text (logical order) → joined letters, in logical order */
export function joinArabic(text) {
  const s = Array.from(text.replace(HARAKAT, "").replace(/[ی]/g, "ي").replace(/[ک]/g, "ك"));
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (!F[c]) { out += c; continue; }
    const prev = i > 0 ? s[i - 1] : "";
    const joinPrev = joinsNext(prev);
    if (c === "ل" && LAM_ALEF[s[i + 1]]) {              // lam + alef → one ligature
      out += String.fromCodePoint(LAM_ALEF[s[i + 1]][joinPrev ? 1 : 0]); i++; continue;
    }
    const next = s[i + 1];
    const joinNext = joinsNext(c) && F[next];
    const forms = F[c];
    const cp = joinPrev && joinNext ? forms[3] : joinPrev ? forms[1] : joinNext ? forms[2] : forms[0];
    out += String.fromCodePoint(cp ?? forms[0]);
  }
  return out;
}
/** joined letters reversed for left-to-right drawing (numbers / Latin inside stay readable) */
export function visualArabic(text) {
  const joined = joinArabic(text);
  const parts = joined.match(/[0-9A-Za-z][0-9A-Za-z.,/:\-]*|./gu) || [];
  return parts.reverse().join("");
}
