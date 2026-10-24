/**
 * Judul bab/subbab tidak membawa tebal sebagai format langsung di run.
 * Word menyalin format langsung judul ke Daftar Isi (entri jadi tebal setelah F9).
 * Tebal cukup berasal dari style judul (Heading 1–3, Judul Bab).
 */
import type JSZip from "jszip";
import { flatten, type DocModel } from "./document-model";

const HEAD_STYLE_RE = /<w:pStyle w:val="(?:Heading[1-3]|JudulBab)"\/>/;
const BOLD_RE = /<w:b(?:Cs)?(?:\s[^>]*)?\/>/g;

export function stripHeadingDirectBold(model: DocModel): number {
  let count = 0;
  for (const n of flatten(model.body)) {
    if (n.kind !== "paragraph") continue;
    const end = n.xml.indexOf("</w:pPr>");
    if (end === -1 || !HEAD_STYLE_RE.test(n.xml.slice(0, end))) continue;
    const runs = n.xml.slice(end);
    const next = runs.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, (rPr) => rPr.replace(BOLD_RE, ""));
    if (next !== runs) { n.xml = n.xml.slice(0, end) + next; count++; }
  }
  return count;
}

/** Pastikan style Heading 1–3 dan Judul Bab memang tebal (sumber tebal satu-satunya). */
export async function ensureBoldHeadingStyles(zip: JSZip): Promise<void> {
  const f = zip.file("word/styles.xml");
  if (!f) return;
  const xml = await f.async("string");
  const out = xml.replace(/<w:style\b[^>]*w:styleId="(?:Heading[1-3]|JudulBab)"[^>]*>[\s\S]*?<\/w:style>/g, (st) => {
    if (/<w:rPr>/.test(st)) {
      return st.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/, (_m, inner: string) => {
        const clean = inner.replace(BOLD_RE, "");
        const lead = /^((?:<w:rStyle\b[^>]*\/>)?(?:<w:rFonts\b[^>]*\/>)?)/.exec(clean)?.[1] ?? "";
        return `<w:rPr>${lead}<w:b/><w:bCs/>${clean.slice(lead.length)}</w:rPr>`;
      });
    }
    return st.replace(/<\/w:style>$/, `<w:rPr><w:b/><w:bCs/></w:rPr></w:style>`);
  });
  if (out !== xml) zip.file("word/styles.xml", out);
}