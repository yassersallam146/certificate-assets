/* Draws a finished PDF on the phone:
 *   background PDF (the form with the placeholders removed)
 * + field map (where each {{PLACEHOLDER}} was, its size, colour, alignment)
 * + the values (client, prices, names …), stamp picture, QR code.
 * Needs window.PDFLib and window.fontkit.
 */
import { joinArabic } from "./arabic.js?v=4";
const AR = /[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]/;
const AR_RUN = /([؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]+(?:[ ،؛؟]+[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]+)*)/;
let fontBytes = {};
async function bytes(url) {
  if (!fontBytes[url]) fontBytes[url] = fetch(url).then(r => { if (!r.ok) throw new Error("font " + url); return r.arrayBuffer(); });
  return fontBytes[url];
}

function hex(c) { const n = parseInt(c.slice(1), 16); return window.PDFLib.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255); }

class Fonts {
  constructor(doc) { this.doc = doc; this.cache = {}; }
  async std(st) {
    const S = window.PDFLib.StandardFonts;
    const key = `${st.serif}|${st.bold}|${st.italic}`;
    if (!this.cache[key]) {
      const name = st.serif
        ? (st.bold && st.italic ? S.TimesRomanBoldItalic : st.bold ? S.TimesRomanBold : st.italic ? S.TimesRomanItalic : S.TimesRoman)
        : (st.bold && st.italic ? S.HelveticaBoldOblique : st.bold ? S.HelveticaBold : st.italic ? S.HelveticaOblique : S.Helvetica);
      this.cache[key] = this.doc.embedFont(name);
    }
    return this.cache[key];
  }
  async arabic(bold) {
    const k = "ar" + bold;
    if (!this.cache[k]) this.cache[k] = bytes(`fonts/NotoNaskhArabic-${bold ? 700 : 400}.ttf`).then(b => this.doc.embedFont(b, { subset: false }));   // full font: keeps joined Arabic letters
    return this.cache[k];
  }
  async uni() {
    if (!this.cache.uni) this.cache.uni = bytes("fonts/DejaVuSans.ttf").then(b => this.doc.embedFont(b, { subset: true }));
    return this.cache.uni;
  }
}

/* split a text into drawable pieces: Arabic runs (Arabic font, right-to-left)
   and other runs (Latin font); in an Arabic sentence the pieces are drawn in reverse order */
async function pieces(fonts, text, st) {
  const out = [];
  const parts = String(text).split(AR_RUN).filter(p => p !== "");
  for (const p of parts) {
    let font, t = p;
    if (AR.test(p)) { font = await fonts.arabic(st.bold); t = joinArabic(p); }
    else {
      font = await fonts.std(st);
      try { font.encodeText(p); } catch { font = await fonts.uni(); }
    }
    out.push({ text: t, font });
  }
  const firstStrong = String(text).match(/[A-Za-z؀-ۿ]/);
  if (firstStrong && AR.test(firstStrong[0])) out.reverse();
  return out;
}

const width = (ps, size) => ps.reduce((a, p) => a + p.font.widthOfTextAtSize(p.text, size), 0);

async function segments(fonts, field, value) {
  const segs = [];
  for (const p of await pieces(fonts, value, field)) segs.push({ ...p, st: field });
  if (field.suffix) {
    const sst = field.suffixStyle || field;
    for (const p of await pieces(fonts, field.suffix, sst)) segs.push({ ...p, st: sst });
  }
  return segs;
}

function wrapWords(text, maxW, font, size) {
  const words = String(text).split(/\s+/); const lines = []; let cur = "";
  for (const w of words) {
    const t = cur ? cur + " " + w : w;
    if (font.widthOfTextAtSize(t, size) <= maxW || !cur) cur = t; else { lines.push(cur); cur = w; }
  }
  if (cur) lines.push(cur);
  return lines;
}

async function drawField(page, fonts, f, value, H, startX) {
  if (value === undefined || value === null) return null;
  if (startX !== undefined && f.align === "left") f = { ...f, x0: startX };
  value = String(value);
  let size = f.size;
  let segs = await segments(fonts, f, value);
  const avail = f.align === "center"
    ? (f.inCell ? f.maxX - f.minX : 2 * Math.min(f.cx - f.minX, f.maxX - f.cx))
    : f.align === "right" ? f.cx - f.minX : f.maxX - f.x0;
  let w = width(segs, size);
  if (w > avail && avail > 10) {                    // shrink a little first
    const s2 = Math.max(f.size * 0.7, size * avail / w);
    size = s2; w = width(segs, size);
  }
  const lines = [];
  if (w > avail + 0.5 && !AR.test(value) && avail > 10) {      // still too long → wrap (Latin text)
    const font = segs[0].font;
    for (const l of wrapWords(value + (f.suffix || ""), avail, font, size)) lines.push(await segments(fonts, { ...f, suffix: "" }, l));
  } else lines.push(segs);
  let base = f.base, endX = f.x0;
  for (const ls of lines) {
    const lw = width(ls, size);
    let x = f.align === "center" ? f.cx - lw / 2 : f.align === "right" ? f.cx - lw : f.x0;
    for (const s of ls) {
      page.drawText(s.text, { x, y: H - base, size, font: s.font, color: hex(s.st.color || "#000000") });
      x += s.font.widthOfTextAtSize(s.text, size);
    }
    endX = x;
    base += size * 1.18;
  }
  return endX;
}

/**
 * bg: ArrayBuffer of the background PDF, layout: JSON, values: {KEY: text}
 * extra: { stamp: PNG bytes, qr: PNG bytes, rows: [{label, desc, price, qty, total}] }
 */
export async function drawPdf(bg, layout, values, extra = {}) {
  const { PDFDocument } = window.PDFLib;
  const doc = await PDFDocument.load(bg);
  doc.registerFontkit(window.fontkit);
  const fonts = new Fonts(doc);
  const pages = doc.getPages();
  const stampImg = extra.stamp ? await doc.embedPng(extra.stamp) : null;
  const qrImg = extra.qr ? await doc.embedPng(extra.qr) : null;
  // fields that follow each other on one line ("{{DATE}} – {{CITY}}") flow one after another
  const order = [...layout.fields].sort((a, b) => a.page - b.page || a.base - b.base || a.x0 - b.x0);
  let prev = null, prevEnd = null;
  for (const f of order) {
    const page = pages[f.page]; const H = page.getHeight();
    const chained = prev && prev.page === f.page && Math.abs(prev.base - f.base) < 1.5 && prev.end !== undefined
      && f.x0 - prev.end < 8 && f.x0 >= prev.x0 && prevEnd !== null && f.align === "left" && prev.align === "left";
    if (f.key === "STAMP_HERE") {
      if (stampImg) {
        const w = 93.6, h = w * stampImg.height / stampImg.width;
        page.drawImage(stampImg, { x: f.x0, y: H - f.y0 - h * 0.78, width: w, height: h });
        stampImg._drawn = (stampImg._drawn || 0) + 1;
      }
      continue;
    }
    const end = await drawField(page, fonts, f, values[f.key], H, chained ? prevEnd : undefined);
    prev = f; prevEnd = end;
    if (f.key === "SERIAL" && qrImg && layout.kind === "certificate") {
      const s = 44.6;
      page.drawImage(qrImg, { x: f.x0 - 78 - s, y: H - f.base - 6, width: s, height: s });
    }
  }
  // package rows
  if (layout.slots && extra.rows) {
    for (let i = 0; i < Math.min(extra.rows.length, layout.slots.length); i++) {
      const slot = layout.slots[i], r = extra.rows[i];
      const page = pages[slot.page]; const H = page.getHeight();
      if (slot.num) await drawField(page, fonts, { ...slot.num, align: "center", base: slot.base, x0: slot.num.cx - 10,
        minX: slot.num.cx - 14, maxX: slot.num.cx + 14, inCell: true, suffix: "" }, String(i + 1), H);
      const c = slot.cols;
      if (c.NAME) await drawField(page, fonts, { ...c.NAME, bold: true }, r.label, H);
      if (c.DESC) await drawField(page, fonts, c.DESC, r.desc, H);
      if (c.PRICE) await drawField(page, fonts, { ...c.PRICE, align: "center", cx: (c.PRICE.minX + c.PRICE.maxX) / 2, inCell: true }, r.price, H);
      if (c.QTY) await drawField(page, fonts, { ...c.QTY, align: "center", cx: (c.QTY.minX + c.QTY.maxX) / 2, inCell: true }, r.qty, H);
      if (c.TOTAL) await drawField(page, fonts, { ...c.TOTAL, align: "center", cx: (c.TOTAL.minX + c.TOTAL.maxX) / 2, inCell: true, bold: true }, r.total, H);
    }
  }
  return await doc.save();
}
