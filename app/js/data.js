/* Device lists, forms and logos from GitHub – kept on the phone for offline use. */
import { GITHUB_RAW } from "./config.js";

const CACHE = "eagle-data-v1";
const LEGACY_CERT = ["AgainPro_template.docx", "Again_template.docx", "Qs4_template.docx", "Replay_Template.docx",
  "SmartPico_template.docx", "SmartxideTouch_Template.docx", "Studio_template.docx", "Toro_template.docx"];

async function cached(path, version, kind) {
  const url = GITHUB_RAW + path + (version ? `?v=${version}` : `?t=${Math.floor(Date.now() / 300000)}`);
  const cache = "caches" in window ? await caches.open(CACHE) : null;
  const key = GITHUB_RAW + path;                     // one cached copy per file
  try {
    const r = await fetch(url, { cache: "no-store" });
    if (!r.ok) throw new Error(`${r.status}`);
    if (cache) await cache.put(key, r.clone());
    return kind === "json" ? await r.json() : await r.arrayBuffer();
  } catch (e) {
    const hit = cache && (await cache.match(key));
    if (hit) return kind === "json" ? await hit.json() : await hit.arrayBuffer();
    throw new Error(kind === "json" ? "No connection and no saved device list yet." :
      `The form ${path.split("/").pop()} is not available (no connection and not saved on this phone).`);
  }
}

export const loadList = name => cached(name, null, "json");

export async function lists() {
  const [quot, framed, unframed, pkg] = await Promise.all([
    loadList("quotation_devices.json"), loadList("devices.json"),
    loadList("devices_unframed.json"), loadList("package_devices.json").catch(() => []),
  ]);
  return { quot, framed, unframed, pkg };
}

export function certTemplateName(e) {
  if (e.template) return e.template;
  const want = e.id.toLowerCase();
  return LEGACY_CERT.find(f => f.split("_")[0].toLowerCase() === want) || `${e.id}_template.docx`;
}

export const quotationForm = e => cached(`templates/quotation/${e.template || e.id + "_quotation.docx"}`, e.v, "bin");
export const certificateForm = (e, framed) => cached(`templates/${framed ? "framed" : "unframed"}/${certTemplateName(e)}`, e.v, "bin");
export const packageForm = () => cached("templates/quotation/package_quotation_template.docx", null, "bin");
export const stampImage = () => cached("logos/stamp.png", null, "bin");
export const logoUrl = GITHUB_RAW + "logos/logo.png";
