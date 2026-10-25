/**
 * Judul tingkat 4 ("a. Rancangan …", "1) Uji prasyarat", "3.4.2.1 Judul", Heading 4–9).
 *
 * Dikenali (hanya di isi bab) bila:
 *  - ber-style Heading 4–9 / outline level 4+, atau bernomor empat tingkat (3.4.2.1), atau
 *  - diawali penanda (a. / a) / (a) / 1. / 1) / A. / ii.) atau penomoran Word, DAN
 *      · teksnya berbentuk judul (≤ 12 kata, diawali kapital, tanpa tanda baca akhir, bukan kalimat),
 *      · langsung diikuti paragraf penjelasan / tabel / gambar (bukan butir lain),
 *      · konteksnya konsisten: dicetak tebal, atau SEMUA butir sejenis di subbab itu berbentuk sama (≥ 2).
 * Daftar biasa ("Bagi mahasiswa:", "Variabel bebas (X): …", butir beruntun) tidak tersentuh.
 *
 * Format: penanda di margin kiri + teks menjorok 0,75 cm (indentasi gantung), tebal, spasi tunggal,
 * menempel ke paragraf berikutnya. Penomoran Word diganti huruf/angka ketikan yang dimulai ulang
 * per subbab. Tidak masuk Daftar Isi (outline level 4).
 */
import type JSZip from "jszip";
import { escapeXml, withProps, PPR_ORDER } from "./xml";
import { formatParagraph, type ApplyReport } from "./apply-format";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { isHeadingRole } from "../structure/roles";
import { leadingRaw, removeLeading } from "./lists";
import { paragraphLine, type VerifyRegion } from "./verify";

const MINOR_STYLE_RE = /^(?:heading|judul|kop)\s*[4-9]$/i;
/** Penanda di awal paragraf (pada aliran run, termasuk tab). */
const MARK_RAW_RE = /^[ \t\u00a0]*(\(?(?:\d{1,2}|[a-zA-Z]|[ivx]{2,4}|[IVX]{2,4})[.)])[ \t\u00a0]*(?=\S)/;
/** Penanda + sisa teks (pada teks paragraf). */
const MARK_TEXT_RE = /^(\(?(?:\d{1,2}|[a-zA-Z]|[ivx]{2,4}|[IVX]{2,4})[.)])\s*(\S.*)$/;
const NUM4_RE = /^\s*\d+\.\d+\.\d+\.\d+\.?\s+(\S.*)$/;

const norm = (s: string) => s.replace(/\s+/g, " ").trim();

export function isMinorHeading(p: ParaNode): boolean {
  if (p.styleId && MINOR_STYLE_RE.test(p.styleId)) return true;
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.xml)?.[0] ?? "";
  const lvl = /<w:outlineLvl w:val="(\d)"/.exec(pPr)?.[1];
  return lvl !== undefined && Number(lvl) >= 3 && Number(lvl) <= 8;
}

/** Teks berbentuk judul, bukan kalimat. */
function titleLike(t: string): boolean {
  const s = norm(t);
  return s.length >= 3 && s.length <= 110 && s.split(" ").length <= 12 &&
    !/[.;:,!?]$/.test(s) && /^[A-Z0-9(\u201c"]/.test(s) && !/\s(adalah|yaitu|ialah|merupakan)\s/i.test(s);
}

function familyOf(marker: string): string {
  const tok = marker.replace(/[().]/g, "");
  const kind = /^\d+$/.test(tok) ? "num" : /^[ivx]{2,4}$/.test(tok) ? "roman" : /^[IVX]{2,4}$/.test(tok) ? "ROMAN"
    : /^[a-z]$/.test(tok) ? "lower" : "upper";
  return `${kind}|${marker.startsWith("(") ? "(" : ""}|${marker.endsWith(")") ? ")" : "."}`;
}

/* ---- numbering.xml (penomoran Word) ---- */

type LvlInfo = { fmt: string; text: string; start: number };

async function readNumbering(zip: JSZip): Promise<(numId: string, ilvl: number) => LvlInfo | null> {
  const xml = (await zip.file("word/numbering.xml")?.async("string")) ?? "";
  const abs = new Map<string, string>();
  for (const m of xml.matchAll(/<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"[^>]*>[\s\S]*?<\/w:abstractNum>/g)) abs.set(m[1], m[0]);
  const num = new Map<string, string>();
  for (const m of xml.matchAll(/<w:num\b[^>]*w:numId="(\d+)"[^>]*>([\s\S]*?)<\/w:num>/g)) {
    const a = /<w:abstractNumId w:val="(\d+)"/.exec(m[2])?.[1];
    if (a) num.set(m[1], a);
  }
  return (numId, ilvl) => {
    const a = num.get(numId);
    const absXml = a ? abs.get(a) : undefined;
    if (!absXml) return null;
    const lvl = new RegExp(`<w:lvl\\b[^>]*w:ilvl="${ilvl}"[^>]*>[\\s\\S]*?<\\/w:lvl>`).exec(absXml)?.[0];
    if (!lvl) return null;
    return {
      fmt: /<w:numFmt w:val="(\w+)"/.exec(lvl)?.[1] ?? "decimal",
      text: /<w:lvlText w:val="([^"]*)"/.exec(lvl)?.[1] ?? `%${ilvl + 1}.`,
      start: Number(/<w:start w:val="(\d+)"/.exec(lvl)?.[1] ?? 1),
    };
  };
}

function roman(n: number): string {
  const map: [number, string][] = [[10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
  let out = "";
  for (const [v, s] of map) while (n >= v) { out += s; n -= v; }
  return out;
}
function formatValue(n: number, fmt: string): string | null {
  switch (fmt) {
    case "decimal": return String(n);
    case "lowerLetter": return String.fromCharCode(96 + n);
    case "upperLetter": return String.fromCharCode(64 + n);
    case "lowerRoman": return roman(n);
    case "upperRoman": return roman(n).toUpperCase();
    default: return null;
  }
}

/* ---- deteksi ---- */

type Cand = {
  p: ParaNode;
  kind: "styled" | "numeric" | "typed" | "native";
  family: string;
  seg: number;
  bold: boolean;
  native?: { numId: string; ilvl: number };
};

export type MinorHeadings = {
  ids: Set<number>;
  count: number;
  /** Dipanggil SETELAH applyFormat/applyListLayout: format judul tingkat 4 diterapkan. */
  apply: (o: { rep: ApplyReport; font: string; halfPts: number; line?: number }) => VerifyRegion[];
};

export async function detectMinorHeadings(
  zip: JSZip,
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
): Promise<MinorHeadings> {
  const lvlOf = await readNumbering(zip);
  const blank = (n: ParaNode | TableNode) => n.kind === "paragraph" && !n.text.trim() && !n.hasDrawing;
  const nextContent = (i: number) => {
    for (let k = i + 1; k < nodes.length; k++) if (!blank(nodes[k])) return k;
    return -1;
  };
  /** Setelahnya ada penjelasan (alinea/tabel/gambar), bukan butir lain. */
  const followedByContent = (i: number) => {
    const k = nextContent(i);
    if (k === -1) return false;
    const m = nodes[k];
    if (m.kind === "table") return true;
    const r = outline.items[k]?.role;
    if (r === "table_caption" || r === "figure" || r === "figure_caption") return true;
    if (r !== "body" || m.hasNumbering || MARK_TEXT_RE.test(norm(m.text)) || isMinorHeading(m)) return false;
    const t = norm(m.text);
    // Alinea penjelasan: panjang, kalimat utuh, atau kalimat pengantar rumus/daftar ("… dengan model:").
    return t.length >= 60 || /[.!?]$/.test(t) || (/:$/.test(t) && t.split(" ").length >= 5);
  };
  const directNum = (p: ParaNode) => {
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.xml)?.[0] ?? "";
    const np = /<w:numPr>([\s\S]*?)<\/w:numPr>/.exec(pPr)?.[1];
    const numId = np ? /<w:numId w:val="(\d+)"/.exec(np)?.[1] : undefined;
    if (!numId || numId === "0") return null;
    return { numId, ilvl: Number(/<w:ilvl w:val="(\d+)"/.exec(np!)?.[1] ?? 0) };
  };

  const cands: Cand[] = [];
  const markerCount = new Map<string, number>();   // butir sejenis per subbab (judul maupun bukan)
  let seg = 0;
  nodes.forEach((n, i) => {
    const it = outline.items[i];
    if (it && isHeadingRole(it.role)) { seg++; return; }
    if (n.kind !== "paragraph" || !it || it.zone !== "main" || it.role !== "body" || n.hasDrawing) return;
    const t = norm(n.text);
    if (!t) return;

    if (isMinorHeading(n)) { cands.push({ p: n, kind: "styled", family: "styled", seg, bold: true }); return; }
    const n4 = NUM4_RE.exec(t);
    if (n4 && titleLike(n4[1])) { cands.push({ p: n, kind: "numeric", family: "num4", seg, bold: true }); return; }

    const native = n.hasNumbering ? directNum(n) : null;
    const mt = n.hasNumbering ? null : MARK_TEXT_RE.exec(t);
    const family = native ? `w:${native.numId}:${native.ilvl}` : mt ? familyOf(mt[1]) : null;
    if (!family) return;
    const key = `${seg}|${family}`;
    markerCount.set(key, (markerCount.get(key) ?? 0) + 1);

    const title = native ? t : mt![2];
    if (!titleLike(title) || !followedByContent(i)) return;
    if (native && !lvlOf(native.numId, native.ilvl)) return;
    if (native && lvlOf(native.numId, native.ilvl)!.fmt === "bullet") return;
    cands.push({ p: n, kind: native ? "native" : "typed", family, seg, bold: n.allBold, native: native ?? undefined });
  });

  // Konsistensi: semua butir sejenis di subbab itu harus berbentuk judul; satu butir saja → wajib tebal.
  const accepted = cands.filter((c) => {
    if (c.kind === "styled" || c.kind === "numeric") return true;
    const key = `${c.seg}|${c.family}`;
    const group = cands.filter((x) => x.seg === c.seg && x.family === c.family).length;
    if (group !== markerCount.get(key)) return false;
    return group >= 2 || c.bold;
  });

  const saved = new Map<ParaNode, string>(accepted.map((c) => [c.p, c.p.xml]));

  const apply: MinorHeadings["apply"] = ({ rep, font, halfPts, line = 240 }) => {
    const regions: VerifyRegion[] = [];
    const order = new Map<string, number>();
    for (const c of accepted) {
      const p = c.p;
      const original = saved.get(p)!;
      p.xml = original;
      const before = paragraphLine(original);

      // Penomoran Word → penanda ketikan, dimulai ulang per subbab.
      let nativeMarker: string | null = null;
      if (c.kind === "native" && c.native) {
        const info = lvlOf(c.native.numId, c.native.ilvl)!;
        const key = `${c.seg}|${c.family}`;
        const idx = (order.get(key) ?? 0) + 1;
        order.set(key, idx);
        const v = formatValue(info.start + idx - 1, info.fmt);
        if (v) nativeMarker = info.text.replace(/%\d/, v);
      }

      // Seluruh judul dimiringkan = gaya judul, bukan istilah asing → dinormalkan.
      const textRuns = [...original.matchAll(/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g)].filter((m) => /<w:t[\s>]/.test(m[1]));
      const allItalic = textRuns.length > 0 && textRuns.every((m) => /<w:i(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(m[1]));

      formatParagraph(p, {
        styleId: c.kind === "styled" ? undefined : null,
        align: "left", line, left: 0, right: 0, firstLine: 0,
        before: line > 240 ? 0 : 180, after: line > 240 ? 0 : 120,
        keepNext: true, keepLines: true,
      }, {
        font, size: halfPts, bold: true, caps: false, color: null, underline: false,
        ...(allItalic ? { italic: false } : {}),
      }, null, rep);

      const rPr = /<w:r(?:\s[^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(p.xml)?.[1] ?? "";
      const insertAt = (x: string) => { const e = x.indexOf("</w:pPr>"); return e !== -1 ? e + 8 : x.indexOf(">") + 1; };
      let marker: string | null = null;
      if (nativeMarker) {
        p.xml = withProps(p.xml, "w:pPr", { "w:numPr": `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>` }, PPR_ORDER);
        p.hasNumbering = false;
        const at = insertAt(p.xml);
        p.xml = p.xml.slice(0, at) + `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(nativeMarker)}</w:t><w:tab/></w:r>` + p.xml.slice(at);
        p.text = `${nativeMarker} ${norm(p.text)}`;
        marker = nativeMarker;
      } else if (c.kind === "typed" || c.kind === "styled") {
        const m = MARK_RAW_RE.exec(leadingRaw(p.xml, 40));
        if (m && !(c.kind === "styled" && p.hasNumbering)) {
          p.xml = removeLeading(p.xml, m[0].length);
          const at = insertAt(p.xml);
          p.xml = p.xml.slice(0, at) + `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(m[1])}</w:t><w:tab/></w:r>` + p.xml.slice(at);
          marker = m[1];
        }
      }

      // Penanda di margin kiri, teks judul (dan baris lanjutannya) sejajar 0,75 cm / 1 cm.
      const props: Record<string, string | null> = {};
      if (marker) {
        const hang = marker.length > 2 ? 567 : 425;
        props["w:tabs"] = `<w:tabs><w:tab w:val="left" w:pos="${hang}"/></w:tabs>`;
        props["w:ind"] = `<w:ind w:left="${hang}" w:right="0" w:hanging="${hang}"/>`;
      } else if (!p.hasNumbering) {
        props["w:tabs"] = null;
        props["w:ind"] = `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`;
      }
      if (c.kind !== "styled") props["w:outlineLvl"] = `<w:outlineLvl w:val="3"/>`;
      p.xml = withProps(p.xml, "w:pPr", props, PPR_ORDER);

      const after = paragraphLine(p.xml);
      if (after !== before) regions.push({ beforeLines: [before].filter(Boolean), afterLines: [after].filter(Boolean), beforeText: "", afterText: "" });
    }
    return regions;
  };

  return { ids: new Set(accepted.map((c) => c.p.id)), count: accepted.length, apply };
}