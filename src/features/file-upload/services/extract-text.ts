import mammoth from "mammoth";
import { extractText, getDocumentProxy } from "unpdf";

export type Extracted = { paragraphs: string[]; kind: "docx" | "pdf" };

export async function extractParagraphs(file: File): Promise<Extracted> {
  const buf = Buffer.from(await file.arrayBuffer());
  if (file.name.toLowerCase().endsWith(".pdf")) {
    const pdf = await getDocumentProxy(new Uint8Array(buf));
    const { text } = await extractText(pdf, { mergePages: true });
    return { kind: "pdf", paragraphs: rebuildPdf(text) };
  }
  const { value } = await mammoth.extractRawText({ buffer: buf });
  return { kind: "docx", paragraphs: value.split(/\n+/).map((s) => s.trim()).filter(Boolean) };
}

function rebuildPdf(text: string): string[] {
  const out: string[] = [];
  let cur = "", lastLen = 0;
  for (const raw of text.split("\n")) {
    const l = raw.trim();
    if (!l || /^\d{1,3}$/.test(l)) { if (cur && !l) { out.push(cur); cur = ""; } continue; }
    const startsNew = /^(bab\s+[ivxlc\d]+|\d+(\.\d+)+\s|gambar\s+\d|tabel\s+\d)/i.test(l);
    if (cur && (startsNew || (lastLen < 70 && /^[A-Z0-9]/.test(l)))) { out.push(cur); cur = ""; }
    cur = cur ? `${cur} ${l}` : l;
    lastLen = l.length;
  }
  if (cur) out.push(cur);
  return out;
}