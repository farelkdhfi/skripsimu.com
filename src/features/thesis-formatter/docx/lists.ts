/**
 * Butir daftar ketikan ("1.", "a)", "(1)", "ii.", "•") → numbering Word asli.
 * Penanda ketikan dibuang dan diganti w:numPr; angka awal, gaya kurung, dan
 * tingkat (1. → a) → i.) dipertahankan.
 */
import JSZip from "jszip";
import { escapeXml, withProps, PPR_ORDER } from "./xml";
import { addContentTypeOverride, findOrAddRelationship } from "./package-parts";
import type { ParaNode } from "./document-model";

type Kind = "decimal" | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman" | "bullet";
type Marker = { kind: Kind; family: string; value: number; lvlText: string; raw: string; token: string };

const ORDERED_RE = /^[ \t\u00a0]*(\()?(\d{1,2}|[a-zA-Z]{1,5})([.)])[ \t\u00a0]+/;
const BULLET_RE = /^[ \t\u00a0]*([\u2022\u25CF\u25BA\u2713\u25AA\u2023\u2043*\-\u2013])[ \t\u00a0]+/;
const ROMAN_RE = /^(x{0,3})(ix|iv|v?i{0,3})$/i;
const MARK_TEXT_RE = /^\s*(?:\(?(?:\d{1,2}|[a-zA-Z]{1,5})[.)]|[\u2022\u25CF\u25BA\u2713\u25AA\u2023\u2043*\-\u2013])\s+/;
const HANG = 357;

const decode = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

function roman(s: string): number {
  const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100 };
  const a = s.toLowerCase();
  let t = 0;
  for (let k = 0; k < a.length; k++) { const c = v[a[k]]; const n = v[a[k + 1]] ?? 0; t += c < n ? -c : c; }
  return t;
}

function parseMarker(s: string): Marker | null {
  const b = BULLET_RE.exec(s);
  if (b) return { kind: "bullet", family: `bullet|${b[1]}`, value: 1, lvlText: b[1], raw: b[0], token: b[1] };
  const m = ORDERED_RE.exec(s);
  if (!m) return null;
  const [raw, open = "", tok, close] = m;
  if (open && close !== ")") return null;
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
  return { kind, family: `${kind}|${open}|${close}`, value, lvlText: `${open}%L${close}`, raw, token: tok };
}

/** Awal teks paragraf (dengan tab & spasi asli) untuk dicocokkan dengan penanda. */
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

function lvlXml(lvl: number, mk: Marker, start: number, font: string, base: number): string {
  const f = escapeXml(font);
  const text = mk.kind === "bullet" ? mk.lvlText : mk.lvlText.replace("%L", `%${lvl + 1}`);
  return `<w:lvl w:ilvl="${lvl}"><w:start w:val="${start}"/><w:numFmt w:val="${mk.kind}"/>` +
    `<w:lvlText w:val="${escapeXml(text)}"/><w:lvlJc w:val="left"/>` +
    `<w:pPr><w:ind w:left="${base + HANG * (lvl + 1)}" w:hanging="${HANG}"/></w:pPr>` +
    `<w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/></w:rPr></w:lvl>`;
}

/** @returns jumlah paragraf yang dikonversi */
export async function convertTypedLists(zip: JSZip, runs: ParaNode[][], font: string, base: number): Promise<number> {
  type Item = { p: ParaNode; mk: Marker; lvl: number };
  const lists: Item[][] = [];

  for (const run of runs) {
    const items = run
      .map((p) => ({ p, mk: parseMarker(leadingRaw(p.xml)) }))
      .filter((x): x is { p: ParaNode; mk: Marker } => x.mk !== null);
    if (!items.length) continue;

    // "i." / "v." tunggal di dalam daftar romawi adalah romawi, bukan huruf.
    for (const kind of ["lowerRoman", "upperRoman"] as const) {
      if (!items.some((x) => x.mk.kind === kind)) continue;
      const letter = kind === "lowerRoman" ? "lowerLetter" : "upperLetter";
      for (const x of items) {
        if (x.mk.kind === letter && /^[ivxIVX]$/.test(x.mk.token)) {
          x.mk = { ...x.mk, kind, family: x.mk.family.replace(letter, kind), value: roman(x.mk.token) };
        }
      }
    }

    const stack: string[] = [];
    lists.push(items.map((x) => {
      let lvl = stack.indexOf(x.mk.family);
      if (lvl === -1) { stack.push(x.mk.family); lvl = stack.length - 1; } else stack.length = lvl + 1;
      return { ...x, lvl: Math.min(lvl, 8) };
    }));
  }
  if (!lists.length) return 0;

  let xml = await zip.file("word/numbering.xml")?.async("string");
  if (!xml) {
    xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>`;
    await addContentTypeOverride(zip, "/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml");
    await findOrAddRelationship(zip, "numbering.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
  }

  let absId = Math.max(-1, ...[...xml.matchAll(/<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"/g)].map((m) => Number(m[1]))) + 1;
  let numId = Math.max(0, ...[...xml.matchAll(/<w:num\b[^>]*w:numId="(\d+)"/g)].map((m) => Number(m[1]))) + 1;
  const abs: string[] = [];
  const nums: string[] = [];
  let count = 0;

  for (const list of lists) {
    const levels = new Map<number, Marker>();
    for (const x of list) if (!levels.has(x.lvl)) levels.set(x.lvl, x.mk);
    const lvls = [...levels.entries()].sort((a, b) => a[0] - b[0])
      .map(([l, mk]) => lvlXml(l, mk, mk.value, font, base)).join("");
    abs.push(`<w:abstractNum w:abstractNumId="${absId}"><w:multiLevelType w:val="multilevel"/>${lvls}</w:abstractNum>`);
    nums.push(`<w:num w:numId="${numId}"><w:abstractNumId w:val="${absId}"/></w:num>`);

    for (const x of list) {
      x.p.xml = withProps(removeLeading(x.p.xml, x.mk.raw.length), "w:pPr", {
        "w:numPr": `<w:numPr><w:ilvl w:val="${x.lvl}"/><w:numId w:val="${numId}"/></w:numPr>`,
        "w:ind": null,
      }, PPR_ORDER);
      x.p.hasNumbering = true;
      x.p.text = x.p.text.replace(MARK_TEXT_RE, "");
      count++;
    }
    absId++;
    numId++;
  }

  const tail = (s: string) => { const m = s.search(/<w:numIdMacAtCleanup\b/); return m !== -1 ? m : s.lastIndexOf("</w:numbering>"); };
  const firstNum = xml.search(/<w:num\b[^>]*w:numId=/);
  const aPos = firstNum !== -1 ? firstNum : tail(xml);
  xml = xml.slice(0, aPos) + abs.join("") + xml.slice(aPos);
  const nPos = tail(xml);
  xml = xml.slice(0, nPos) + nums.join("") + xml.slice(nPos);
  zip.file("word/numbering.xml", xml);
  return count;
}