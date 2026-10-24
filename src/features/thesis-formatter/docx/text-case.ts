/**
 * Teks yang hanya TAMPIL kapital (w:caps) diubah menjadi kapital sungguhan.
 * Tanpa ini "Bab III" tetap tersimpan huruf kecil: muncul apa adanya saat
 * Daftar Isi diperbarui Word, saat di-copy, atau di aplikasi lain.
 */
import { escapeXml } from "./xml";
import { flatten, type DocModel } from "./document-model";

const CAPS_ON = /<w:caps(?:\s+w:val="(?:1|true|on)")?\s*\/>/;
const decodeXml = (s: string) =>
  s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

export function realizeCaps(model: DocModel): number {
  let count = 0;
  for (const p of flatten(model.body)) {
    if (p.kind !== "paragraph" || !p.xml.includes("<w:caps")) continue;
    let all = true;
    let changed = false;
    const xml = p.xml.replace(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g, (run) => {
      if (!/<w:t[\s>]/.test(run)) return run;
      const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(run)?.[0] ?? "";
      if (!CAPS_ON.test(rPr)) {
        if (/<w:t(?:\s[^>]*)?>[^<]*\S/.test(run)) all = false;
        return run;
      }
      return run.replace(/(<w:t(?:\s[^>]*)?>)([\s\S]*?)(<\/w:t>)/g, (_m, a: string, t: string, c: string) => {
        const raw = decodeXml(t);
        const up = raw.toUpperCase();
        if (up !== raw) changed = true;
        return a + escapeXml(up) + c;
      });
    });
    if (!changed) continue;
    p.xml = xml;
    if (all) p.text = p.text.toUpperCase();
    count++;
  }
  return count;
}