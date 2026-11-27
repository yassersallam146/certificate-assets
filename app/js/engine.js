/* Eagle Generator – document engine (browser)
 * Fills the same Word forms as the Windows app ({{PLACEHOLDERS}}),
 * builds package-quotation rows, adds the stamp picture and the QR code.
 * Needs window.JSZip.
 */
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const IMG_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/* ── number → words (SAR), same as the Windows app ─────────────────────── */
const ONES = ["", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven",
  "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"];
const TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"];
function numText(n) {
  if (n === 0) return "";
  if (n < 20) return ONES[n];
  if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? " " + ONES[n % 10] : "");
  if (n < 1000) return ONES[Math.floor(n / 100)] + " hundred" + (n % 100 ? " " + numText(n % 100) : "");
  if (n < 1e6) return numText(Math.floor(n / 1000)) + " thousand" + (n % 1000 ? " " + numText(n % 1000) : "");
  if (n < 1e9) return numText(Math.floor(n / 1e6)) + " million" + (n % 1e6 ? " " + numText(n % 1e6) : "");
  if (n < 1e12) return numText(Math.floor(n / 1e9)) + " billion" + (n % 1e9 ? " " + numText(n % 1e9) : "");
  return String(n);
}
export function wordsSAR(n) {
  const t = numText(Math.floor(n)).trim();
  return t ? t[0].toUpperCase() + t.slice(1) + " SAR only" : "Zero SAR only";
}

/* ── amounts as people type them (250000 · 250,000 · 250.000 · Arabic digits) ── */
export function parseAmount(v) {
  let s = String(v ?? "").trim().replace(/(sar|sr|ريال|ر\.س)/gi, "").replace(/[\s ]/g, "")
    .replace(/٬/g, ",").replace(/٫/g, ".").replace(/[٠-٩]/g, d => "٠١٢٣٤٥٦٧٨٩".indexOf(d));
  if (!s) return 0;
  if (!/^[\d.,]+$/.test(s)) throw new Error(`'${v}' is not a valid amount.`);
  if (s.includes(",") && s.includes(".")) {
    const dec = s.lastIndexOf(",") > s.lastIndexOf(".") ? "," : ".";
    s = s.split(dec === "," ? "." : ",").join("").replace(dec, ".");
  } else if (s.includes(",") || s.includes(".")) {
    const sep = s.includes(",") ? "," : ".";
    const parts = s.split(sep);
    s = (parts.length > 2 || parts.slice(1).every(p => p.length === 3)) ? parts.join("") : s.replace(",", ".");
  }
  return parseFloat(s);
}
export function parseQty(v) {
  const s = String(v ?? "").trim().replace(/[٠-٩]/g, d => "٠١٢٣٤٥٦٧٨٩".indexOf(d)) || "1";
  if (!/^\d+$/.test(s) || +s < 1) throw new Error(`Quantity '${v}' must be a whole number.`);
  return +s;
}
export const fmt = n => Math.floor(n).toLocaleString("en-US");
export function longDate(iso) {          // "2026-11-27" → "Friday, 27 November 2026"
  const d = new Date(iso + "T12:00:00");
  if (isNaN(d)) return iso;
  const day = d.toLocaleDateString("en-GB", { weekday: "long" });
  const rest = d.toLocaleDateString("en-GB", { day: "2-digit", month: "long", year: "numeric" });
  return `${day}, ${rest}`;
}
export function shortDate(iso) {         // "2026-11-27" → "27/11/2026"
  const [y, m, d] = iso.split("-");
  return d ? `${d}/${m}/${y}` : iso;
}
export function safeName(t) {
  return String(t || "").replace(/[\\/:*?"<>|\x00-\x1f]/g, "-").trim().replace(/\s+/g, "_")
    .replace(/^[._-]+|[._-]+$/g, "").slice(0, 80) || "file";
}

/* ── XML helpers ─────────────────────────────────────────────────────── */
const parse = s => new DOMParser().parseFromString(s, "application/xml");
const ser = d => new XMLSerializer().serializeToString(d);
const byTag = (node, tag) => Array.from(node.getElementsByTagNameNS(W_NS, tag));

function parts(zip) {
  return Object.keys(zip.files).filter(n =>
    /^word\/(document|header\d*|footer\d*)\.xml$/.test(n));
}

/* Replace every key of map in one XML document (run-level first, then merged
 * paragraph text – handles placeholders split over several runs). */
function replaceInDoc(doc, map) {
  const keys = Object.keys(map);
  for (const p of byTag(doc, "p")) {
    // only the paragraph's own runs (text boxes are separate paragraphs)
    const ts = byTag(p, "t").filter(t => closestP(t) === p);
    if (!ts.length) continue;
    let full = ts.map(t => t.textContent).join("");
    if (!keys.some(k => full.includes(k))) continue;
    for (const t of ts) {
      let x = t.textContent;
      for (const k of keys) if (x.includes(k)) x = x.split(k).join(map[k]);
      if (x !== t.textContent) { t.textContent = x; t.setAttribute("xml:space", "preserve"); }
    }
    full = ts.map(t => t.textContent).join("");
    if (keys.some(k => full.includes(k))) {
      for (const k of keys) full = full.split(k).join(map[k]);
      ts[0].textContent = full; ts[0].setAttribute("xml:space", "preserve");
      ts.slice(1).forEach(t => (t.textContent = ""));
    }
  }
}
function closestP(node) {
  let n = node.parentNode;
  while (n && !(n.namespaceURI === W_NS && n.localName === "p")) n = n.parentNode;
  return n;
}

export async function openDocx(arrayBuffer) {
  return await window.JSZip.loadAsync(arrayBuffer);
}
export async function fill(zip, map) {
  for (const name of parts(zip)) {
    const doc = parse(await zip.file(name).async("string"));
    replaceInDoc(doc, map);
    zip.file(name, ser(doc));
  }
}
export async function textOf(zip) {
  let out = "";
  for (const name of parts(zip)) {
    const doc = parse(await zip.file(name).async("string"));
    out += byTag(doc, "t").map(t => t.textContent).join("") + "\n";
  }
  return out;
}

/* ── pictures ─────────────────────────────────────────────────────────── */
let picCounter = 0;
async function addImage(zip, partName, bytes, ext = "png") {
  const media = `media/eagle_${Date.now()}_${picCounter++}.${ext}`;
  zip.file("word/" + media, bytes);
  const relsName = partName.replace("word/", "word/_rels/") + ".rels";
  let rels = zip.file(relsName) ? parse(await zip.file(relsName).async("string"))
    : parse(`<Relationships xmlns="${PKG_NS}"/>`);
  const root = rels.documentElement;
  const ids = Array.from(root.getElementsByTagName("Relationship")).map(r => r.getAttribute("Id"));
  let n = 1000; while (ids.includes("rIdE" + n)) n++;
  const rel = rels.createElementNS(PKG_NS, "Relationship");
  rel.setAttribute("Id", "rIdE" + n); rel.setAttribute("Type", IMG_REL); rel.setAttribute("Target", media);
  root.appendChild(rel);
  zip.file(relsName, ser(rels));
  const ct = parse(await zip.file("[Content_Types].xml").async("string"));
  const has = Array.from(ct.documentElement.getElementsByTagName("Default"))
    .some(d => d.getAttribute("Extension").toLowerCase() === ext);
  if (!has) {
    const d = ct.createElementNS(ct.documentElement.namespaceURI, "Default");
    d.setAttribute("Extension", ext); d.setAttribute("ContentType", ext === "png" ? "image/png" : "image/jpeg");
    ct.documentElement.appendChild(d);
    zip.file("[Content_Types].xml", ser(ct));
  }
  return "rIdE" + n;
}
function drawingRunXml(rId, cx, cy, id) {
  return `<w:r xmlns:w="${W_NS}"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="${cx}" cy="${cy}"/>
 <wp:docPr id="${id}" name="EaglePic${id}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
 <pic:nvPicPr><pic:cNvPr id="${id}" name="EaglePic${id}"/><pic:cNvPicPr/></pic:nvPicPr>
 <pic:blipFill><a:blip r:embed="${rId}" xmlns:r="${R_NS}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
 <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
 </pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}
async function imageSize(bytes) {
  const blob = new Blob([bytes]);
  const bmp = await createImageBitmap(blob);
  const s = { w: bmp.width, h: bmp.height }; bmp.close?.(); return s;
}

/* {{STAMP_HERE}} → stamp picture (1.3 inch wide), like "Stamp Word Files" */
export async function stamp(zip, stampBytes) {
  const { w, h } = await imageSize(stampBytes);
  const cx = Math.round(1.3 * 914400), cy = Math.round(cx * h / w);
  let count = 0;
  for (const name of parts(zip)) {
    const doc = parse(await zip.file(name).async("string"));
    const paras = byTag(doc, "p").filter(p => byTag(p, "t").filter(t => closestP(t) === p)
      .map(t => t.textContent).join("").includes("{{STAMP_HERE}}"));
    if (!paras.length) continue;
    const rId = await addImage(zip, name, stampBytes, "png");
    for (const p of paras) {
      const ts = byTag(p, "t").filter(t => closestP(t) === p);
      const full = ts.map(t => t.textContent).join("").replace("{{STAMP_HERE}}", "");
      ts.forEach((t, i) => (t.textContent = i ? "" : full));
      const run = parse(drawingRunXml(rId, cx, cy, 9000 + count)).documentElement;
      p.appendChild(doc.importNode(run, true));
      count++;
    }
    zip.file(name, ser(doc));
  }
  return count;
}

/* QR code at the start of the footer line ("Certificate No: …") */
export async function addFooterQR(zip, pngBytes) {
  const footers = Object.keys(zip.files).filter(n => /^word\/footer\d*\.xml$/.test(n));
  for (const name of footers) {
    const doc = parse(await zip.file(name).async("string"));
    const p = byTag(doc, "p").find(p => byTag(p, "t").some(t => t.textContent.trim()));
    if (!p) continue;
    const rId = await addImage(zip, name, pngBytes, "png");
    const size = Math.round(0.62 * 914400);
    const run = doc.importNode(parse(drawingRunXml(rId, size, size, 9500)).documentElement, true);
    const firstRun = byTag(p, "r")[0];
    firstRun ? p.insertBefore(run, firstRun) : p.appendChild(run);
    const sp = doc.importNode(parse(`<w:r xmlns:w="${W_NS}"><w:t xml:space="preserve">  </w:t></w:r>`).documentElement, true);
    p.insertBefore(sp, run.nextSibling);
    zip.file(name, ser(doc));
    return true;
  }
  return false;
}
export function qrPng(text) {
  const qr = window.qrcode(0, "M");
  qr.addData(unescape(encodeURIComponent(text)));
  qr.make();
  const n = qr.getModuleCount(), cell = 6, pad = 6, size = n * cell + pad * 2;
  const c = document.createElement("canvas"); c.width = c.height = size;
  const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, size, size); g.fillStyle = "#000";
  for (let r = 0; r < n; r++) for (let k = 0; k < n; k++) if (qr.isDark(r, k)) g.fillRect(pad + k * cell, pad + r * cell, cell, cell);
  return new Promise(res => c.toBlob(b => b.arrayBuffer().then(res), "image/png"));
}

/* ── Package quotation: one table row per selected device ─────────────── */
const ROW_RE = /\{\{[^{}]+_(NAME|DESC|PRICE|QTY|TOTAL)\}\}/;
function cellXml(txt, bold, size, color, align, fill) {
  const esc = String(txt).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return `<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${fill}"/></w:tcPr><w:p><w:pPr><w:jc w:val="${align}"/></w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>${bold ? "<w:b/>" : ""}<w:sz w:val="${size * 2}"/><w:color w:val="${color}"/></w:rPr>
<w:t xml:space="preserve">${esc}</w:t></w:r></w:p></w:tc>`;
}
export async function packageRows(zip, rows /* [{label,desc,price,qty}] */, grandTotal) {
  const doc = parse(await zip.file("word/document.xml").async("string"));
  const rowText = tr => byTag(tr, "t").map(t => t.textContent).join(" ");
  const tables = byTag(doc, "tbl");
  const tbl = tables.find(t => byTag(t, "tr").some(tr => /\{\{[^{}]+_NAME\}\}/.test(rowText(tr))));
  if (!tbl) return false;
  const trs = byTag(tbl, "tr").filter(tr => tr.parentNode === tbl);
  const kill = trs.filter(tr => ROW_RE.test(rowText(tr)));
  const totals = ["Grand Total", "TOTAL_PRICE", "Amount in Words", "Down Payment", "Remaining", "Installments", fmt(grandTotal)];
  const anchor = trs.find(tr => !kill.includes(tr) && totals.some(x => rowText(tr).includes(x))) || null;
  kill.forEach(tr => tbl.removeChild(tr));
  rows.forEach((r, i) => {
    const fill = i % 2 ? "EBF3FB" : "FFFFFF";
    const xml = `<w:tr xmlns:w="${W_NS}">` +
      cellXml(i + 1, false, 12, "333333", "center", fill) + cellXml(r.label, true, 12, "1F4E79", "left", fill) +
      cellXml(r.desc, false, 11, "444444", "left", fill) + cellXml(fmt(r.price), false, 12, "333333", "center", fill) +
      cellXml(r.qty, false, 12, "333333", "center", fill) + cellXml(fmt(r.price * r.qty), true, 12, "1F4E79", "center", fill) +
      `</w:tr>`;
    const node = doc.importNode(parse(xml).documentElement, true);
    anchor ? tbl.insertBefore(node, anchor) : tbl.appendChild(node);
  });
  zip.file("word/document.xml", ser(doc));
  return true;
}

export async function toBlob(zip) {
  return await zip.generateAsync({ type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

/* ── high level ───────────────────────────────────────────────────────── */
export function quotationMap(f, total, price, qty, dp) {
  return {
    "{{DEVICE}}": f.deviceLabel || f.device || "", "{{ACCOUNT}}": f.account, "{{CITY}}": f.city,
    "{{EMP_CODE}}": f.empCode, "{{PRICE}}": fmt(price), "{{QTY}}": String(qty),
    "{{TOTAL_PRICE}}": fmt(total), "{{TOTAL_PRICE_WORDS}}": wordsSAR(total),
    "{{DOWN_PAYMENT}}": `${Math.floor(dp)}%`, "{{REMAINING_PERCENT}}": `${Math.floor(100 - dp)}%`,
    "{{INSTALLMENTS}}": f.installments || "", "{{DATE}}": longDate(f.date), "{{NOTES}}": f.notes || "",
    "{{SERIAL}}": f.serial, "{{EMP_NAME}}": f.empName, "{{EMP_PHONE}}": f.empPhone,
  };
}
