/**
 * Judul bab diseragamkan jadi SATU paragraf judul dua baris (dijalankan sebelum outline):
 *   "BAB I" ↵ "PENDAHULUAN"       — label bab + judul di dua paragraf → digabung
 *   "I" ↵ "PENDAHULUAN"           — kampus yang hanya memakai angka romawi (tanpa kata BAB)
 *   "I. PENDAHULUAN"              — romawi + judul satu baris → dipecah dua baris
 * Konvensi dokumen dipertahankan (romawi saja tidak ditambah "BAB").
 * Romawi tanpa kata BAB hanya diterima bila berurutan I, II, III, … di seluruh dokumen,
 * judulnya kapital, dan label/judulnya ditonjolkan (rata tengah / tebal / Heading 1),
 * supaya butir "I." biasa tidak ikut.
 * Judul bab satu baris "BAB II. TINJAUAN PUSTAKA" dipecah oleh normalizeHeadings.
 */
import { escapeXml, withProps, PPR_ORDER } from "./xml";
import type { DocModel, ParaNode } from "./document-model";
import { leadingRaw, removeLeading } from "./lists";
import { paragraphLine, type VerifyRegion } from "./verify";

const WORD_LABEL_RE = /^(bab|chapter)\s+([ivxlcdm]+|\d{1,2})\s*[.:]?$/i;
const BARE_LABEL_RE = /^([IVX]{1,5})\s*\.?$/;
const BARE_LINE_TEXT_RE = /^([IVX]{1,5})\s*\.?\s+(\S.*)$/;
const BARE_LINE_RAW_RE = /^[ \t\u00a0]*([IVX]{1,5})[ \t\u00a0]*\.?[ \t\u00a0]+/;
const PAGE_TAIL_RE = /(\.{3,}|\u2026|\t)\s*[\divxlcdm]+\s*$/i;

const norm = (s: string) => s.replace(/\s+/g, " ").trim();

function fromRoman(s: string): number {
  const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
  const a = s.toLowerCase();
  let t = 0;
  for (let k = 0; k < a.length; k++) { const c = v[a[k]] ?? 0; const n = v[a[k + 1]] ?? 0; t += c < n ? -c : c; }
  return t;
}

function upperish(t: string): boolean {
  const letters = t.replace(/[^A-Za-z]/g, "");
  if (letters.length < 3) return false;
  return letters.replace(/[^A-Z]/g, "").length / letters.length >= 0.8;
}

function titleText(t: string): boolean {
  return t.length >= 3 && t.length <= 120 && !/[.;:,!?]$/.test(t) && !/^\d/.test(t) && !/^[a-z]/.test(t) &&
    t.split(" ").length <= 16 && !WORD_LABEL_RE.test(t) && !BARE_LABEL_RE.test(t) && !PAGE_TAIL_RE.test(t);
}

const emphasized = (p: ParaNode) =>
  p.align === "center" || p.allBold || /^(?:heading|judul)\s*1$|^title$/i.test(p.styleId ?? "");

export type ChapterHeadReport = {
  merged: number;
  split: number;
  /** Paragraf bab bernomor romawi saja (tanpa kata BAB) — dipakai outline sebagai penanda bab. */
  chapterIds: Set<number>;
  /** Paragraf judul yang digabung ke paragraf labelnya (dibuang dari body). */
  removed: Set<ParaNode>;
  regions: VerifyRegion[];
};

export function unifyChapterHeadings(model: DocModel): ChapterHeadReport {
  const body = model.body;
  const out: ChapterHeadReport = { merged: 0, split: 0, chapterIds: new Set(), removed: new Set(), regions: [] };
  const isEmpty = (n: DocModel["body"][number]) =>
    n.kind === "paragraph" && !n.text.trim() && !n.hasDrawing && !n.hasSectPr && !n.hasPageBreak && !n.hasField;
  const nextText = (i: number) => {
    for (let k = i + 1; k < body.length && k <= i + 4; k++) {
      const n = body[k];
      if (isEmpty(n)) continue;
      return n.kind === "paragraph" ? k : -1;
    }
    return -1;
  };

  type Cand = { i: number; pair: boolean; bare: boolean; value: number; title: number };
  const cands: Cand[] = [];
  for (let i = 0; i < body.length; i++) {
    const p = body[i];
    if (p.kind !== "paragraph" || p.hasDrawing || p.hasNumbering) continue;
    const t = norm(p.text);
    if (!t || t.length > 120) continue;

    const w = WORD_LABEL_RE.exec(t);
    const b = w ? null : BARE_LABEL_RE.exec(t);
    if (w || b) {
      const k = nextText(i);
      if (k === -1) continue;
      const q = body[k] as ParaNode;
      if (q.hasDrawing || q.hasNumbering || q.hasSectPr || !titleText(norm(q.text))) continue;
      if (b && !(upperish(q.text) && (emphasized(p) || emphasized(q)))) continue;
      cands.push({ i, pair: true, bare: !w, value: w ? 0 : fromRoman(b![1]), title: k });
      continue;
    }

    const l = BARE_LINE_TEXT_RE.exec(t);
    if (l && upperish(l[2]) && titleText(l[2]) && emphasized(p) && !p.hasSectPr) {
      cands.push({ i, pair: false, bare: true, value: fromRoman(l[1]), title: -1 });
    }
  }

  // Romawi tanpa kata BAB: hanya bila berurutan I, II, III, … (minimal dua bab).
  const bare = cands.filter((c) => c.bare);
  const bareOk = bare.length >= 2 && bare.every((c, k) => c.value === k + 1);
  const accepted = cands.filter((c) => !c.bare || bareOk).sort((a, b) => b.i - a.i);

  for (const c of accepted) {
    const p = body[c.i] as ParaNode;
    const beforeLines = [paragraphLine(p.xml)];
    if (c.pair) {
      const q = body[c.title] as ParaNode;
      beforeLines.push(paragraphLine(q.xml));
      const qs = q.xml.indexOf("</w:pPr>");
      const content = q.xml.slice(qs !== -1 ? qs + 8 : q.xml.indexOf(">") + 1, q.xml.length - "</w:p>".length);
      let xml = p.xml.replace(/<\/w:p>$/, `<w:r><w:br/></w:r>${content}</w:p>`);
      // Label tanpa style mengambil style judulnya (mis. Heading 1).
      const qStyle = qs !== -1 ? /<w:pStyle w:val="[^"]+"\/>/.exec(q.xml.slice(0, qs))?.[0] : undefined;
      const pEnd = xml.indexOf("</w:pPr>");
      if (qStyle && !/<w:pStyle\b/.test(pEnd !== -1 ? xml.slice(0, pEnd) : "")) {
        xml = withProps(xml, "w:pPr", { "w:pStyle": qStyle }, PPR_ORDER);
      }
      p.xml = xml;
      p.text = `${norm(p.text)} ${norm(q.text)}`;
      for (let k = c.title; k > c.i; k--) { out.removed.add(body[k] as ParaNode); body.splice(k, 1); }
      out.merged++;
    } else {
      const m = BARE_LINE_RAW_RE.exec(leadingRaw(p.xml, 20));
      if (!m) continue;
      const rPr = /<w:r(?:\s[^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(p.xml)?.[1] ?? "";
      let xml = removeLeading(p.xml, m[0].length);
      const e = xml.indexOf("</w:pPr>");
      const at = e !== -1 ? e + 8 : xml.indexOf(">") + 1;
      xml = xml.slice(0, at) + `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(m[1])}</w:t><w:br/></w:r>` + xml.slice(at);
      p.xml = xml;
      p.text = `${m[1]} ${norm(p.text).replace(/^[IVX]{1,5}\s*\.?\s+/, "")}`;
      out.split++;
    }
    if (c.bare) out.chapterIds.add(p.id);
    out.regions.push({ beforeLines: beforeLines.filter(Boolean), afterLines: [paragraphLine(p.xml)].filter(Boolean), beforeText: "", afterText: "" });
  }
  return out;
}