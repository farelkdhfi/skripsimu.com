/**
 * Alur teks & halaman di isi skripsi.
 *  1. Pindah halaman manual di tengah bab dibuang (tetap bila sebelum judul pembuka halaman).
 *  2. Baris kosong (Enter tanpa isi) yang hanya jadi jarak dibuang — jarak sudah diatur
 *     formatter. Yang bermakna dibiarkan (pemisah dua tabel, ruang tanda tangan, dsb.).
 *  3. Paragraf isi dirapikan: ganti baris manual → spasi, spasi ganda → satu,
 *     indentasi manual (spasi/tab di awal) dibuang, baris rumus rata tengah.
 *  4. Paragraf yang terbelah digabung kembali.
 */
import type JSZip from "jszip";
import { withProps, PPR_ORDER } from "./xml";
import type { DocModel, ParaNode, TableNode } from "./document-model";
import { isEmptyParagraph, nextVisible, prevVisible } from "./document-model";
import type { Outline } from "../structure/outline";
import { isHeadingRole, startsNewPage, type Role } from "../structure/roles";
import { paragraphLine, type VerifyRegion } from "./verify";
import { leadingRaw, removeLeading } from "./lists";
import { editText } from "./captions";

const PAGE_BR = /<w:br\b[^>]*w:type="page"[^>]*\/>/g;
const SOFT_BR_G = /<w:br(?![^>]*w:type="(?:page|column)")[^>]*\/>|<w:cr\s*\/>/g;
const SOFT_BR = /<w:br(?![^>]*w:type="(?:page|column)")[^>]*\/>|<w:cr\s*\/>/;
const WIDOW_OFF = /<w:widowControl\s+w:val="(?:0|false|off)"\s*\/>/g;

const decodeXml = (s: string) =>
  s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const fullText = (xml: string) =>
  [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => decodeXml(m[1])).join("");

/**
 * Baris rumus: ada "=", pendek, hampir tanpa kata biasa ("Y = α + β1·X1 + ε"). Diperiksa pada
 * teks ASLI (dengan tab): baris bertab ("i ⇥ = Perlakuan ke-i") dan baris "simbol = Arti" adalah
 * keterangan rumus, bukan rumus.
 */
const isFormula = (raw: string) => {
  if (/\t/.test(raw)) return false;
  const t = raw.replace(/\s+/g, " ").trim();
  if (/^\S{1,12}\s*=\s+[A-Z][a-z]{3,}/.test(t)) return false;
  return /=/.test(t) && t.length <= 160 && (t.match(/[A-Za-z\u00C0-\u024F]{4,}/g) ?? []).length <= 2;
};

export type FlowReport = {
  breaks: number;
  blanks: number;
  softBreaks: number;
  spaces: number;
  formulas: number;
  merged: number;
  regions: VerifyRegion[];
};

export function cleanFlow(
  model: DocModel,
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  exclude: Set<number>,
): FlowReport {
  const out: FlowReport = { breaks: 0, blanks: 0, softBreaks: 0, spaces: 0, formulas: 0, merged: 0, regions: [] };
  const roleOf = new Map<number, Role>();
  const zoneOf = new Map<number, string>();
  nodes.forEach((n, i) => {
    const it = outline.items[i];
    if (it) { roleOf.set(n.id, it.role); zoneOf.set(n.id, it.zone); }
  });
  const body = model.body;
  type Node = DocModel["body"][number];
  const isPara = (n: Node | undefined): n is ParaNode => !!n && n.kind === "paragraph";
  const roleAt = (n: Node | undefined) => (isPara(n) ? roleOf.get(n.id) ?? "body" : undefined);
  const isCaption = (n: Node | undefined) => roleAt(n) === "table_caption" || roleAt(n) === "figure_caption";
  const hasContent = (n: Node) =>
    n.kind === "table" || (isPara(n) && (Boolean(n.text.trim()) || n.hasDrawing));

  /* ---- 1. Pindah halaman yang tidak semestinya ---- */
  for (let k = 0; k < body.length; k++) {
    const p = body[k];
    if (!isPara(p) || zoneOf.get(p.id) !== "main") continue;

    if (/<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(p.xml) && !startsNewPage(roleAt(p)!)) {
      p.xml = p.xml.replace(/<w:pageBreakBefore\b[^>]*\/>/, "");
      out.breaks++;
    }
    if (!/<w:br\b[^>]*w:type="page"/.test(p.xml)) continue;

    let j = k + 1;
    while (j < body.length && !hasContent(body[j])) j++;
    const next = body[j];
    if (isPara(next) && startsNewPage(roleAt(next)!)) continue;

    p.xml = p.xml.replace(PAGE_BR, "");
    p.hasPageBreak = false;
    out.breaks++;
    if (isEmptyParagraph(p) && !p.hasDrawing && !p.hasSectPr) {
      body.splice(k, 1);
      k--;
    }
  }

  /* ---- 2. Baris kosong ---- */
  const isBlank = (n: Node | undefined): n is ParaNode =>
    isPara(n) && isEmptyParagraph(n) && !n.hasDrawing && !n.hasSectPr;
  const anchored = (n: Node | undefined) =>
    !!n && (n.kind === "table" || isCaption(n) || isHeadingRole(roleAt(n) ?? "body") || (isPara(n) && exclude.has(n.id)));

  for (let k = 0; k < body.length; k++) {
    if (!isBlank(body[k])) continue;
    let e = k;
    while (e + 1 < body.length && isBlank(body[e + 1])) e++;
    // Tetangga yang TAMPAK (bookmark di antaranya tidak dihitung): mis. dua tabel yang hanya
    // dipisah baris kosong + bookmark tetap harus dipisah.
    const prev = body[prevVisible(body, k)];
    const next = body[nextVisible(body, e)];
    const zone = zoneOf.get((body[k] as ParaNode).id);

    let from = -1;
    if (next && zone === "main") from = k;                              // isi bab: semua dibuang
    else if (next && zone === "back" && e === k && (anchored(prev) || anchored(next))) from = k;

    if (from !== -1 && prev?.kind === "table" && next?.kind === "table") from = k + 1; // jangan satukan dua tabel
    if (from === -1 || from > e) { k = e; continue; }
    out.blanks += e - from + 1;
    body.splice(from, e - from + 1);
    k = from - 1;
  }

  /* ---- 3. Paragraf isi ---- */
  for (const n of body) {
    if (!isPara(n)) continue;
    if (WIDOW_OFF.test(n.xml)) { WIDOW_OFF.lastIndex = 0; n.xml = n.xml.replace(WIDOW_OFF, ""); }
    WIDOW_OFF.lastIndex = 0;
    if (zoneOf.get(n.id) !== "main" || roleAt(n) !== "body" || exclude.has(n.id) || n.hasDrawing || !n.text.trim()) continue;

    const before = paragraphLine(n.xml);
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(n.xml)?.[0] ?? "";
    const t = n.text.replace(/\s+/g, " ").trim();

    if (isFormula(n.text)) {
      const indented = /<w:ind\b[^>]*w:(?:firstLine|left|start)="[1-9]/.test(pPr);
      if (!/<w:jc w:val="center"/.test(pPr) || indented) {
        n.xml = withProps(n.xml, "w:pPr", {
          "w:ind": `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`,
          "w:jc": `<w:jc w:val="center"/>`,
        }, PPR_ORDER);
        out.formulas++;
      }
      continue;
    }
    if (/<w:jc w:val="(center|right|end)"/.test(pPr)) continue;

    let changed = false;
    if (t.length >= 60 && SOFT_BR.test(n.xml)) {
      n.xml = n.xml.replace(SOFT_BR_G, `<w:t xml:space="preserve"> </w:t>`);
      n.text = n.text.replace(/\n/g, " ");
      out.softBreaks++;
      changed = true;
    }
    if (!/<w:numPr>/.test(n.xml)) {
      const lead = /^[ \t\u00a0]+/.exec(leadingRaw(n.xml, 40))?.[0];
      if (lead) {
        n.xml = removeLeading(n.xml, lead.length);
        n.text = n.text.replace(/^[ \t\u00a0]+/, "");
        out.spaces++;
        changed = true;
      }
    }
    const full = fullText(n.xml);
    const edits = [...full.matchAll(/ {2,}/g)].map((m) => ({ start: m.index! + 1, end: m.index! + m[0].length, text: "" }));
    if (edits.length) {
      n.xml = editText(n.xml, edits);
      n.text = n.text.replace(/ {2,}/g, " ");
      out.spaces += edits.length;
      changed = true;
    }
    if (changed) {
      const after = paragraphLine(n.xml);
      if (after !== before) {
        out.regions.push({ beforeLines: [before].filter(Boolean), afterLines: [after].filter(Boolean), beforeText: n.text, afterText: n.text });
      }
    }
  }

  /* ---- 4. Paragraf terbelah ---- */
  const vocab = new Set<string>();
  for (const n of body) {
    if (isPara(n)) for (const w of n.text.toLowerCase().match(/[a-z\u00c0-\u024f]{3,}/g) ?? []) vocab.add(w);
  }
  const mergeable = (n: Node | undefined): n is ParaNode =>
    isPara(n) && zoneOf.get(n.id) === "main" && roleAt(n) === "body" &&
    !exclude.has(n.id) && !n.hasDrawing && !n.hasSectPr && Boolean(n.text.trim()) &&
    !/<w:numPr>/.test(n.xml) && !/<w:jc w:val="(center|right|end)"/.test(n.xml) &&
    !/<w:pStyle w:val="Caption"/.test(n.xml) && !/<m:oMath/.test(n.xml) && !isFormula(n.text.trim());
  const runsOf = (xml: string) => {
    const s = xml.indexOf("</w:pPr>");
    const start = s !== -1 ? s + 8 : xml.indexOf(">") + 1;
    return xml.slice(start, xml.lastIndexOf("</w:p>"));
  };

  for (let k = 0; k < body.length - 1; k++) {
    const a = body[k];
    if (!mergeable(a)) continue;
    const b = body[k + 1];
    if (!mergeable(b)) continue;

    const ta = a.text.replace(/\s+$/, "");
    const tb = b.text.replace(/^\s+/, "");
    if (ta.length < 60) continue;
    const lowerNext = /^[a-z\u00DF-\u00FF]/.test(tb) && /[A-Za-z\u00C0-\u024F,]$/.test(ta);
    const digitNext = /^\d/.test(tb) && /[A-Za-z\u00C0-\u024F]$/.test(ta);
    if (!lowerNext && !digitNext) continue;

    const rawGap = /\s$/.test(a.text) || /^\s/.test(b.text);
    const wa = /([A-Za-z\u00C0-\u024F]+)$/.exec(ta)?.[1]?.toLowerCase() ?? "";
    const wb = /^([a-z\u00C0-\u024F]+)/.exec(tb)?.[1] ?? "";
    const glue = rawGap ? "" : wa && wb && vocab.has(wa + wb) ? "" : " ";

    const before = [paragraphLine(a.xml), paragraphLine(b.xml)].filter(Boolean);
    const at = a.xml.lastIndexOf("</w:p>");
    a.xml = a.xml.slice(0, at) +
      (glue ? `<w:r><w:t xml:space="preserve"> </w:t></w:r>` : "") +
      runsOf(b.xml) + "</w:p>";
    a.text = a.text + glue + b.text;
    body.splice(k + 1, 1);
    out.regions.push({ beforeLines: before, afterLines: [paragraphLine(a.xml)].filter(Boolean), beforeText: a.text, afterText: a.text });
    out.merged++;
    k--;
  }
  return out;
}

/** Kontrol baris yatim (widow/orphan) dinyalakan untuk seluruh dokumen. */
export async function patchWidowControl(zip: JSZip): Promise<void> {
  const f = zip.file("word/styles.xml");
  if (!f) return;
  let xml = await f.async("string");
  xml = xml.replace(WIDOW_OFF, "");
  if (/<w:pPrDefault>\s*<w:pPr>/.test(xml)) {
    xml = xml.replace(/<w:pPrDefault>\s*<w:pPr>([\s\S]*?)<\/w:pPr>/, (m, inner: string) =>
      /<w:widowControl\b/.test(inner) ? m : m.replace("<w:pPr>", "<w:pPr><w:widowControl/>"));
  } else if (/<w:pPrDefault\s*\/>|<w:pPrDefault>\s*<\/w:pPrDefault>/.test(xml)) {
    xml = xml.replace(/<w:pPrDefault\s*\/>|<w:pPrDefault>\s*<\/w:pPrDefault>/, "<w:pPrDefault><w:pPr><w:widowControl/></w:pPr></w:pPrDefault>");
  } else if (/<w:docDefaults>/.test(xml)) {
    xml = xml.replace(/<\/w:docDefaults>/, "<w:pPrDefault><w:pPr><w:widowControl/></w:pPr></w:pPrDefault></w:docDefaults>");
  }
  zip.file("word/styles.xml", xml);
}