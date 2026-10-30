/**
 * Daftar (list) di isi skripsi.
 *  1. Daftar ketikan ("1.", "a)", "(1)", "ii.", "•", simbol Symbol/Wingdings) → numbering Word.
 *  2. Daftar Word asli (style List Number/Bullet, numPr) → dirapikan.
 *
 * Aturan seragam:
 *  - nomor tingkat-1 sejajar margin kiri (awal teks subbab); teks butir pada STEP; tiap tingkat +STEP
 *  - pemisah nomor seragam "1." / "a." / "i."; simbol dibiarkan apa adanya
 *  - nomor mulai ulang per subbab & setelah kalimat pengantar ("…sebagai berikut:")
 *  - daftar yang terpotong paragraf penjelasan tetap bersambung (a → penjelasan → b)
 *  - paragraf penjelasan di bawah butir sejajar teks butirnya
 *  - tingkat ditentukan indentasi asli (bukan bentuk simbol)
 *  - lampiran, sampul, dan sel tabel tidak disentuh
 */
import JSZip from "jszip";
import { escapeXml, withProps, PPR_ORDER } from "./xml";
import { addContentTypeOverride, findOrAddRelationship } from "./package-parts";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { isHeadingRole } from "../structure/roles";
import { paragraphLine, type VerifyRegion } from "./verify";

/** Disiapkan untuk template kampus nanti. */
export const LIST_STYLE = {
  /** Jarak nomor → teks dan tambahan tiap tingkat (twip). 425 ≈ 0,75 cm. */
  step: 425,
  /** Pemisah nomor daftar bernomor. */
  separator: "." as "." | ")",
};

type Kind = "decimal" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman" | "bullet";
type Marker = {
  kind: Kind;
  family: string;
  value: number;
  raw: string;
  glyph: string;
  glyphFont: string | null;
  /** Elemen <w:sym> bila penanda berupa simbol font (Symbol/Wingdings). */
  sym: string | null;
};

/* ------------------------------------------------------------------ */
/* Pembacaan teks awal paragraf                                         */
/* ------------------------------------------------------------------ */

const decode = (s: string) =>
  s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

/** Awal teks paragraf (dengan tab & spasi asli) untuk dicocokkan dengan penanda. */
const HEADING_STYLE_ID_RE = /^(?:heading|judul)\s*[1-9]$/i;
function headingStyled(n: ParaNode): boolean {
  const t = n.text.trim();
  if (!t || t.length > 150 || /[.!?;:]$/.test(t)) return false;
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(n.xml)?.[0] ?? "";
  return HEADING_STYLE_ID_RE.test(n.styleId ?? "") || /<w:outlineLvl w:val="[0-8]"/.test(pPr);
}

export function leadingRaw(xml: string, max = 24): string {
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>/g;
  re.lastIndex = Math.max(0, xml.indexOf("</w:pPr>"));
  let out = "";
  let m: RegExpExecArray | null;
  while (out.length < max && (m = re.exec(xml))) out += m[0].startsWith("<w:tab") ? "\t" : decode(m[1]);
  return out;
}

/** Buang `count` karakter pertama (teks + tab) dari aliran run. */
export function removeLeading(xml: string, count: number): string {
  const start = Math.max(0, xml.indexOf("</w:pPr>"));
  const head = xml.slice(0, start);
  let rest = xml.slice(start);
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>/g;
  let left = count;
  let m: RegExpExecArray | null;
  while (left > 0 && (m = re.exec(rest))) {
    if (m[0].startsWith("<w:tab")) {
      rest = rest.slice(0, m.index) + rest.slice(m.index + m[0].length);
      re.lastIndex = m.index;
      left--;
      continue;
    }
    const text = decode(m[1]);
    const take = Math.min(left, text.length);
    const repl = `<w:t xml:space="preserve">${escapeXml(text.slice(take))}</w:t>`;
    rest = rest.slice(0, m.index) + repl + rest.slice(m.index + m[0].length);
    re.lastIndex = m.index + repl.length;
    left -= take;
  }
  return head + rest;
}

/* ------------------------------------------------------------------ */
/* Penanda daftar ketikan                                               */
/* ------------------------------------------------------------------ */

const BULLETS =
  "\\u2022\\u25CF\\u25CB\\u25E6\\u25AA\\u25AB\\u25A0\\u25A1\\u25C6\\u25C7\\u25BA\\u25B6\\u27A2\\u27A4" +
  "\\u2192\\u2713\\u2714\\u2756\\u2023\\u2043\\u00B7*\\-\\u2013\\u2014\\uF0A7\\uF0B7\\uF076\\uF0D8\\uF0FC\\uF0A8\\uF0D7";
const BULLET_RE = new RegExp(`^[ \\t\\u00a0]*([${BULLETS}])[ \\t\\u00a0]+`);
const ORDERED_RE = /^[ \t\u00a0]*(\()?(\d{1,2}|[a-zA-Z]{1,5})([.)])([ \t\u00a0]+|(?=[A-Za-z\u00C0-\u024F(]))/;
const ROMAN_RE = /^(x{0,3})(ix|iv|v?i{0,3})$/i;
/** Kode karakter bullet umum pada font Symbol/Wingdings. */
const SYM_BULLET = new Set(["B7", "A7", "76", "D8", "FC", "A8", "6E", "71", "E8", "D7", "A1", "77"]);

function roman(s: string): number {
  const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100 };
  const a = s.toLowerCase();
  let t = 0;
  for (let k = 0; k < a.length; k++) { const c = v[a[k]]; const n = v[a[k + 1]] ?? 0; t += c < n ? -c : c; }
  return t;
}

const runsStart = (xml: string) => xml.slice(Math.max(0, xml.indexOf("</w:pPr>")));
const firstFont = (xml: string) => /<w:rFonts\b[^>]*w:ascii="([^"]+)"/.exec(runsStart(xml))?.[1] ?? null;

function symMarker(xml: string): { glyph: string; font: string; el: string } | null {
  const rest = runsStart(xml);
  const m = /<w:sym\b[^>]*\/>/.exec(rest);
  if (!m) return null;
  const before = rest.slice(0, m.index);
  if (/<w:t(?:\s[^>]*)?>[^<]*\S/.test(before) || /<w:(drawing|pict)\b/.test(before)) return null;
  const font = /w:font="([^"]+)"/.exec(m[0])?.[1];
  const ch = /w:char="([0-9A-Fa-f]{2,4})"/.exec(m[0])?.[1];
  if (!font || !ch || !/symbol|wingdings|webdings/i.test(font)) return null;
  if (!SYM_BULLET.has(ch.slice(-2).toUpperCase())) return null;
  return { glyph: String.fromCharCode(parseInt(ch.length === 2 ? `F0${ch}` : ch, 16)), font, el: m[0] };
}

export function parseMarker(xml: string): Marker | null {
  const sym = symMarker(xml);
  if (sym) return { kind: "bullet", family: "bullet", value: 1, raw: "", glyph: sym.glyph, glyphFont: sym.font, sym: sym.el };

  const s = leadingRaw(xml);
  const b = BULLET_RE.exec(s);
  if (b) {
    const g = b[1];
    const pua = g.charCodeAt(0) >= 0xf000;
    return { kind: "bullet", family: "bullet", value: 1, raw: b[0], glyph: g, glyphFont: pua ? firstFont(xml) ?? "Symbol" : null, sym: null };
  }

  const m = ORDERED_RE.exec(s);
  if (!m) return null;
  const [raw, open = "", tok, close, gap] = m;
  if (open && close !== ")") return null;
  if (!gap && !(/^\d+$/.test(tok) || close === ")")) return null; // "S.Kom" bukan penanda
  let kind: Kind;
  let value: number;
  if (/^\d+$/.test(tok)) { kind = "decimal"; value = Number(tok); }
  else if (tok.length > 1) {
    if (!ROMAN_RE.test(tok)) return null;
    kind = tok === tok.toLowerCase() ? "lowerRoman" : "upperRoman";
    value = roman(tok);
  } else {
    kind = tok === tok.toLowerCase() ? "lowerLetter" : "upperLetter";
    value = tok.toLowerCase().charCodeAt(0) - 96;
  }
  return { kind, family: `${kind}|${open}|${close}`, value, raw, glyph: "", glyphFont: null, sym: null };
}

function stripMarker(xml: string, mk: Marker): string {
  if (mk.sym) {
    const at = Math.max(0, xml.indexOf("</w:pPr>"));
    const out = xml.slice(0, at) + xml.slice(at).replace(mk.sym, "");
    const ws = /^[ \t\u00a0]*/.exec(leadingRaw(out))![0].length;
    return ws ? removeLeading(out, ws) : out;
  }
  return removeLeading(xml, mk.raw.length);
}

function textAfterMarker(text: string, mk: Marker): string {
  const t = text.replace(/^[ \t\u00a0]+/, "");
  const tok = mk.raw.trim();
  return (tok && t.startsWith(tok) ? t.slice(tok.length) : t).replace(/^[ \t\u00a0]+/, "");
}

/** Posisi horizontal penanda di dokumen asli (twip), untuk menentukan tingkat. */
function markerPos(xml: string, raw: string): number {
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? "";
  const ind = /<w:ind\b[^>]*\/>/.exec(pPr)?.[0] ?? "";
  const a = (n: string) => Number(new RegExp(`w:${n}="(-?\\d+)"`).exec(ind)?.[1] ?? 0);
  let pos = (a("left") || a("start")) + a("firstLine") - a("hanging");
  for (const ch of /^[ \t\u00a0]*/.exec(raw)?.[0] ?? "") pos += ch === "\t" ? 720 : 60;
  return pos;
}

const isHeadingLike = (t: string) =>
  t.length <= 80 && !/[.;:!?,]$/.test(t.trim()) && t.trim().split(/\s+/).length <= 10;

/* ------------------------------------------------------------------ */
/* numbering.xml & styles.xml                                           */
/* ------------------------------------------------------------------ */

type Numbering = { abs: Map<string, string>; num: Map<string, string> };

function parseNumbering(xml: string): Numbering {
  const abs = new Map<string, string>();
  for (const m of xml.matchAll(/<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"[^>]*>[\s\S]*?<\/w:abstractNum>/g)) abs.set(m[1], m[0]);
  const num = new Map<string, string>();
  for (const m of xml.matchAll(/<w:num\b[^>]*w:numId="(\d+)"[^>]*>([\s\S]*?)<\/w:num>/g)) {
    const a = /<w:abstractNumId w:val="(\d+)"/.exec(m[2])?.[1];
    if (a) num.set(m[1], a);
  }
  return { abs, num };
}

type StyleNum = { numId?: string; ilvl?: number; basedOn?: string };

function parseStyleNumbering(xml: string): Map<string, StyleNum> {
  const map = new Map<string, StyleNum>();
  for (const m of xml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    if (!/w:type="paragraph"/.test(m[1])) continue;
    const id = /w:styleId="([^"]+)"/.exec(m[1])?.[1];
    if (!id) continue;
    const np = /<w:numPr>([\s\S]*?)<\/w:numPr>/.exec(m[2])?.[1];
    const ilvl = np ? /<w:ilvl w:val="(\d+)"/.exec(np)?.[1] : undefined;
    map.set(id, {
      numId: np ? /<w:numId w:val="(\d+)"/.exec(np)?.[1] : undefined,
      ilvl: ilvl !== undefined ? Number(ilvl) : undefined,
      basedOn: /<w:basedOn w:val="([^"]+)"/.exec(m[2])?.[1],
    });
  }
  return map;
}

function effectiveNum(p: ParaNode, styles: Map<string, StyleNum>): { numId: string; ilvl: number } | null {
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.xml)?.[0] ?? "";
  const np = /<w:numPr>([\s\S]*?)<\/w:numPr>/.exec(pPr)?.[1];
  let numId = np ? /<w:numId w:val="(\d+)"/.exec(np)?.[1] : undefined;
  let ilvl = np ? /<w:ilvl w:val="(\d+)"/.exec(np)?.[1] : undefined;
  let sid = /<w:pStyle w:val="([^"]+)"/.exec(pPr)?.[1];
  for (let d = 0; sid && d < 6 && (numId === undefined || ilvl === undefined); d++) {
    const s = styles.get(sid);
    if (!s) break;
    if (numId === undefined && s.numId) numId = s.numId;
    if (ilvl === undefined && s.ilvl !== undefined) ilvl = String(s.ilvl);
    sid = s.basedOn;
  }
  if (!numId || numId === "0") return null;
  return { numId, ilvl: Math.min(8, Number(ilvl ?? 0)) };
}

function levelFmt(absXml: string, ilvl: number): string | null {
  const lvl = new RegExp(`<w:lvl\\b[^>]*w:ilvl="${ilvl}"[^>]*>[\\s\\S]*?<\\/w:lvl>`).exec(absXml)?.[0];
  return lvl ? /<w:numFmt w:val="(\w+)"/.exec(lvl)?.[1] ?? "decimal" : null;
}

/* ------------------------------------------------------------------ */
/* Definisi tingkat                                                     */
/* ------------------------------------------------------------------ */

const S = () => LIST_STYLE.step;
const indXml = (lvl: number) => `<w:ind w:left="${(lvl + 1) * S()}" w:hanging="${S()}"/>`;
const paraInd = (lvl: number) => `<w:ind w:left="${(lvl + 1) * S()}" w:right="0" w:hanging="${S()}"/>`;

function typedLvl(lvl: number, mk: Marker, start: number, font: string, keep = false): string {
  // keep: bentuk penanda asli dipertahankan ("(1)", "1)", "a.") — mis. pedoman yang memakai (1), (2).
  const [, open, close] = mk.family.split("|");
  const text = mk.kind === "bullet" ? mk.glyph
    : keep ? `${open ? "(" : ""}%${lvl + 1}${close || "."}` : `%${lvl + 1}${LIST_STYLE.separator}`;
  const f = escapeXml(mk.kind === "bullet" && mk.glyphFont ? mk.glyphFont : font);
  return `<w:lvl w:ilvl="${lvl}"><w:start w:val="${start}"/><w:numFmt w:val="${mk.kind}"/>` +
    `<w:lvlText w:val="${escapeXml(text)}"/><w:lvlJc w:val="left"/>` +
    `<w:pPr>${indXml(lvl)}</w:pPr>` +
    `<w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}" w:hint="default"/></w:rPr></w:lvl>`;
}

/** Salinan abstractNum daftar Word asli dengan geometri & pemisah yang diseragamkan. */
function cloneAbstract(absXml: string, newId: number, keep = false): string | null {
  if (/<w:numStyleLink\b/.test(absXml)) return null;
  let x = absXml
    .replace(/w:abstractNumId="\d+"/, `w:abstractNumId="${newId}"`)
    .replace(/<w:(nsid|tmpl|styleLink|name)\b[^>]*\/>/g, "");
  x = x.replace(/<w:lvl\b[^>]*w:ilvl="(\d)"[^>]*>[\s\S]*?<\/w:lvl>/g, (lvl, k: string) => {
    const fmt = /<w:numFmt w:val="(\w+)"/.exec(lvl)?.[1] ?? "decimal";
    let out = lvl.replace(/<w:pStyle\b[^>]*\/>/, "");
    if (!keep && fmt !== "bullet" && fmt !== "none" && LIST_STYLE.separator === ".") {
      out = out.replace(/<w:lvlText w:val="([^"]*)"\/>/, (_m, t: string) =>
        `<w:lvlText w:val="${t.replace(/\((%\d)\)/g, "$1.").replace(/(%\d)\)/g, "$1.")}"/>`);
    }
    const ind = indXml(Number(k));
    if (/<w:pPr>[\s\S]*?<\/w:pPr>/.test(out)) return out.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, `<w:pPr>${ind}</w:pPr>`);
    if (/<w:rPr>/.test(out)) return out.replace(/<w:rPr>/, `<w:pPr>${ind}</w:pPr><w:rPr>`);
    return out.replace(/<\/w:lvl>$/, `<w:pPr>${ind}</w:pPr></w:lvl>`);
  });
  return x;
}

/* ------------------------------------------------------------------ */
/* Proses utama                                                         */
/* ------------------------------------------------------------------ */

/** Indentasi yang dipasang ulang SETELAH applyFormat (yang menimpa w:ind). */
export type ListLayout = Map<number, { left: number; hanging?: number } | { restore: string | null }>;

export type ListReport = {
  items: number;
  typed: number;
  native: number;
  lists: number;
  layout: ListLayout;
  regions: VerifyRegion[];
};

export async function formatLists(
  zip: JSZip,
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  opts: { font: string; exclude: Set<number>; markers?: "uniform" | "keep" },
): Promise<ListReport> {
  const rep: ListReport = { items: 0, typed: 0, native: 0, lists: 0, layout: new Map(), regions: [] };
  const keep = opts.markers === "keep";

  let numXml = (await zip.file("word/numbering.xml")?.async("string")) ?? null;
  const stylesXml = (await zip.file("word/styles.xml")?.async("string")) ?? "";
  const styleNum = parseStyleNumbering(stylesXml);
  const numbering = numXml ? parseNumbering(numXml) : null;

  let absId = numXml ? Math.max(-1, ...[...numXml.matchAll(/<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"/g)].map((m) => Number(m[1]))) + 1 : 0;
  let numId = numXml ? Math.max(0, ...[...numXml.matchAll(/<w:num\b[^>]*w:numId="(\d+)"/g)].map((m) => Number(m[1]))) + 1 : 1;
  const absOut: string[] = [];
  const numOut: string[] = [];
  const addNum = (abstractXml: string) => {
    absOut.push(abstractXml);
    numOut.push(`<w:num w:numId="${numId}"><w:abstractNumId w:val="${absId}"/></w:num>`);
    rep.lists++;
    absId++;
    return numId++;
  };

  /* ---- 1. Kandidat: penanda ketikan & daftar Word asli ---- */

  type Cand = { kind: "typed"; mk: Marker } | { kind: "native"; numId: string; ilvl: number; dup: Marker | null };
  const cands = new Map<number, Cand>();
  const inScope = (i: number) => {
    const it = outline.items[i];
    return it && (it.zone === "main" || it.zone === "front") && it.role === "body";
  };
  nodes.forEach((n, i) => {
    if (n.kind !== "paragraph") return;
    const it = outline.items[i];
    // Lampiran: tidak dikonversi, tapi indentasi aslinya disimpan karena applyFormat
    // memberi indentasi gantung pada baris berpenanda ("A.", "1.") — merusak formulir/kuesioner.
    if (it?.zone === "back" && !isHeadingRole(it.role) && parseMarker(n.xml)) {
      const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(n.xml)?.[0] ?? "";
      rep.layout.set(n.id, { restore: /<w:ind\b[^>]*\/>/.exec(pPr)?.[0] ?? null });
      return;
    }
    if (!inScope(i) || opts.exclude.has(n.id) || n.hasDrawing) return;
    // Paragraf ber-style judul (Heading/Judul 1–9 atau outline level) yang pendek = judul, bukan
    // butir daftar — walau memakai penomoran otomatis Word.
    if (headingStyled(n)) return;
    const nat = numbering ? effectiveNum(n, styleNum) : null;
    if (nat) {
      const a = numbering!.num.get(nat.numId);
      const fmt = a ? levelFmt(numbering!.abs.get(a) ?? "", nat.ilvl) : null;
      if (fmt && fmt !== "none") { cands.set(i, { kind: "native", ...nat, dup: parseMarker(n.xml) }); return; }
    }
    if (n.hasNumbering) return;
    const mk = parseMarker(n.xml);
    if (mk) cands.set(i, { kind: "typed", mk });
  });

  // Huruf/romawi hanya diterima bila ada tetangga berurutan di subbab yang sama
  // ("A. Comte menyatakan…" bukan butir daftar).
  {
    let seg = 0;
    const segOf = new Map<number, number>();
    nodes.forEach((_n, i) => { if (isHeadingRole(outline.items[i]?.role ?? "body")) seg++; segOf.set(i, seg); });
    const typed = [...cands.entries()].filter(([, c]) => c.kind === "typed") as [number, { kind: "typed"; mk: Marker }][];
    for (const [i, c] of typed) {
      if (c.mk.kind === "decimal" || c.mk.kind === "bullet") continue;
      const ok = typed.some(([j, d]) => j !== i && segOf.get(j) === segOf.get(i) &&
        d.mk.family === c.mk.family && Math.abs(d.mk.value - c.mk.value) === 1);
      if (!ok) cands.delete(i);
    }
  }

  /* ---- 2. Penelusuran dokumen ---- */

  type Inst = { family: string; level: number; numId: number; last: number; pos: number; glyph: string };
  let stack: Inst[] = [];
  const nativeOpen = new Map<string, number>();
  let lastItem: { level: number; headingLike: boolean } | null = null;
  let pending: ParaNode[] = [];

  const own = (level: number) => {
    for (const q of pending) rep.layout.set(q.id, { left: (level + 1) * S() });
    pending = [];
  };
  const closeAll = () => {
    if (lastItem?.headingLike) own(lastItem.level);
    pending = [];
    stack = [];
    nativeOpen.clear();
    lastItem = null;
  };
  const newTyped = (mk: Marker, level: number, fam: string, pos: number): Inst => {
    const id = addNum(`<w:abstractNum w:abstractNumId="${absId}"><w:multiLevelType w:val="multilevel"/>${typedLvl(level, mk, mk.value, opts.font, keep)}</w:abstractNum>`);
    return { family: fam, level, numId: id, last: mk.value, pos, glyph: mk.glyph };
  };
  const place = (p: ParaNode, xml: string, level: number, nid: number, headingLike: boolean) => {
    const before = paragraphLine(p.xml);
    p.xml = withProps(xml, "w:pPr", {
      "w:numPr": `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${nid}"/></w:numPr>`,
      "w:ind": paraInd(level),
    }, PPR_ORDER);
    p.hasNumbering = true;
    rep.layout.set(p.id, { left: (level + 1) * S(), hanging: S() });
    const after = paragraphLine(p.xml);
    if (before !== after) {
      const t = after;
      rep.regions.push({ beforeLines: [before].filter(Boolean), afterLines: [after].filter(Boolean), beforeText: t, afterText: t });
    }
    lastItem = { level, headingLike };
    rep.items++;
  };

  nodes.forEach((n, i) => {
    if (n.kind === "table") return;                       // tabel tidak memutus daftar
    const it = outline.items[i];
    if (!it || (it.zone !== "main" && it.zone !== "front")) { closeAll(); return; }
    if (isHeadingRole(it.role)) { closeAll(); return; }
    if (it.role !== "body" || opts.exclude.has(n.id) || n.hasDrawing || !n.text.trim()) return;

    const c = cands.get(i);

    /* ---- daftar Word asli ---- */
    if (c?.kind === "native") {
      let nid = nativeOpen.get(c.numId);
      if (nid === undefined) {
        const a = numbering!.num.get(c.numId);
        const clone = a ? cloneAbstract(numbering!.abs.get(a)!, absId, keep) : null;
        nid = clone ? addNum(clone) : Number(c.numId);
        nativeOpen.set(c.numId, nid);
        if (lastItem?.headingLike) own(lastItem.level);
        pending = [];
      } else if (lastItem) own(lastItem.level);
      const xml = c.dup ? stripMarker(n.xml, c.dup) : n.xml;
      if (c.dup) n.text = textAfterMarker(n.text, c.dup);
      place(n, xml, c.ilvl, nid, isHeadingLike(n.text));
      rep.native++;
      return;
    }

    /* ---- daftar ketikan ---- */
    if (c?.kind === "typed") {
      const mk = c.mk;
      const pos = markerPos(n.xml, mk.raw);
      const fam = mk.kind === "bullet" ? `bullet@${Math.round(pos / 360)}` : mk.family;
      const separated = pending.length > 0;
      const idx = stack.findIndex((s) => s.family === fam);
      const restart = idx !== -1 && mk.kind !== "bullet" && mk.value === 1 && stack[idx].last >= 1;
      let inst: Inst;

      if (idx !== -1 && !restart) {
        // lanjutan butir yang sama (a → penjelasan → b)
        if (lastItem) own(lastItem.level);
        stack.length = idx + 1;
        inst = stack[idx];
        if (mk.kind === "bullet" && inst.glyph !== mk.glyph) {
          inst = newTyped(mk, inst.level, fam, pos);   // simbol lain, tingkat sama
          stack[idx] = inst;
        }
        inst.last = mk.value;
      } else if (idx > 0) {
        // sub-daftar baru di bawah butir induk berikutnya
        if (lastItem) own(lastItem.level);
        stack.length = idx;
        inst = newTyped(mk, idx, fam, pos);
        stack.push(inst);
      } else {
        const top = stack[0];
        const unrelated = !stack.length || idx === 0 || (separated && pos <= (top?.pos ?? 0) + 150);
        if (unrelated) {
          // daftar baru
          if (lastItem?.headingLike) own(lastItem.level);
          pending = [];
          stack = [];
          inst = newTyped(mk, 0, fam, pos);
        } else {
          // sub-daftar dengan jenis penanda berbeda
          if (lastItem) own(lastItem.level);
          inst = newTyped(mk, Math.min(stack.length, 8), fam, pos);
        }
        stack.push(inst);
      }

      const xml = stripMarker(n.xml, mk);
      n.text = textAfterMarker(n.text, mk);
      place(n, xml, inst.level, inst.numId, isHeadingLike(n.text));
      rep.typed++;
      return;
    }

    /* ---- paragraf biasa ---- */
    if (!lastItem) return;
    pending.push(n);
    // Kalimat pengantar ("…sebagai berikut:") memulai daftar Word asli yang baru.
    if (/[:：]\s*$/.test(n.text.trim())) nativeOpen.clear();
  });
  closeAll();

  /* ---- 3. Tulis numbering.xml ---- */

  if (absOut.length) {
    if (!numXml) {
      numXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>`;
      await addContentTypeOverride(zip, "/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
      await findOrAddRelationship(zip, "numbering.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
    }
    const tail = (s: string) => { const m = s.search(/<w:numIdMacAtCleanup\b/); return m !== -1 ? m : s.lastIndexOf("</w:numbering>"); };
    const firstNum = numXml.search(/<w:num\b[^>]*w:numId=/);
    const aPos = firstNum !== -1 ? firstNum : tail(numXml);
    numXml = numXml.slice(0, aPos) + absOut.join("") + numXml.slice(aPos);
    const nPos = tail(numXml);
    numXml = numXml.slice(0, nPos) + numOut.join("") + numXml.slice(nPos);
    zip.file("word/numbering.xml", numXml);
  }

  return rep;
}

/**
 * Dipanggil SETELAH applyFormat: butir daftar mendapat indentasi gantung,
 * paragraf penjelasan disejajarkan dengan teks butirnya (indentasi baris pertama tetap),
 * baris berpenanda di lampiran dikembalikan ke indentasi aslinya.
 */
export function applyListLayout(nodes: (ParaNode | TableNode)[], layout: ListLayout): number {
  let count = 0;
  for (const p of nodes) {
    if (p.kind !== "paragraph") continue;
    const l = layout.get(p.id);
    if (!l) continue;
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.xml)?.[0] ?? "";

    if ("restore" in l) {
      p.xml = withProps(p.xml, "w:pPr", { "w:ind": l.restore }, PPR_ORDER);
      count++;
      continue;
    }

    let ind: string;
    if (l.hanging !== undefined) {
      ind = `<w:ind w:left="${l.left}" w:right="0" w:hanging="${l.hanging}"/>`;
    } else {
      // Rumus / paragraf rata tengah-kanan tidak digeser (posisinya relatif ke margin).
      if (/<w:jc w:val="(center|right|end)"/.test(pPr) || /<m:oMath/.test(p.xml)) continue;
      const cur = /<w:ind\b[^>]*\/>/.exec(pPr)?.[0] ?? "";
      const fl = /w:firstLine="(\d+)"/.exec(cur)?.[1] ?? "0";
      ind = `<w:ind w:left="${l.left}" w:right="0" w:firstLine="${fl}"/>`;
    }
    p.xml = withProps(p.xml, "w:pPr", { "w:ind": ind }, PPR_ORDER);
    count++;
  }
  return count;
}