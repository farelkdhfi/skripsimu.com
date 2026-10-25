/**
 * Subbab bernomor otomatis Word ("1.1" dari penomoran Heading 2, teks paragrafnya hanya
 * "Latar Belakang") → nomornya ditulis menjadi teks ("1.1⇥Latar Belakang") dan penomoran
 * otomatisnya dimatikan. Setelah itu subbab dikenali seperti subbab ketikan, nomornya ikut
 * diperiksa, dan Daftar Isi memuatnya dengan benar.
 *
 * Nomor dihitung dari konteks bab (bukan dari penghitung numbering Word, yang sering
 * tercampur dengan daftar lain). Hanya penomoran desimal bertingkat ("%1.%2", "%1.%2.%3")
 * yang ditulis ulang. Dijalankan setelah unifyChapterHeadings dan sebelum outline.
 */
import type JSZip from "jszip";
import { escapeXml, withProps, PPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";
import { leadingRaw } from "./lists";
import { paragraphLine, type VerifyRegion } from "./verify";

const HEAD_STYLE_RE = /^(?:heading|judul)\s*([1-3])$/i;
const TYPED_RE = /^\s*(\d+(?:\.\d+){1,3})[.)]?\s/;
const CHAPTER_RE = /^\s*(?:bab|chapter)\s+([ivxlcdm]+|\d+)\b/i;
const BARE_RE = /^\s*([IVX]{1,5})\b/;

function fromRoman(s: string): number {
  const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
  const a = s.toLowerCase();
  let t = 0;
  for (let k = 0; k < a.length; k++) { const c = v[a[k]] ?? 0; const n = v[a[k + 1]] ?? 0; t += c < n ? -c : c; }
  return t;
}

type StyleInfo = { numId?: string; ilvl?: number; basedOn?: string; level?: number };

function readStyles(xml: string): Map<string, StyleInfo> {
  const map = new Map<string, StyleInfo>();
  for (const m of xml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    if (!/w:type="paragraph"/.test(m[1])) continue;
    const id = /w:styleId="([^"]+)"/.exec(m[1])?.[1];
    if (!id) continue;
    const np = /<w:numPr>([\s\S]*?)<\/w:numPr>/.exec(m[2])?.[1];
    const il = np ? /<w:ilvl w:val="(\d+)"/.exec(np)?.[1] : undefined;
    const ol = /<w:outlineLvl w:val="(\d)"/.exec(m[2])?.[1];
    const hs = HEAD_STYLE_RE.exec(id)?.[1];
    map.set(id, {
      numId: np ? /<w:numId w:val="(\d+)"/.exec(np)?.[1] : undefined,
      ilvl: il !== undefined ? Number(il) : undefined,
      basedOn: /<w:basedOn w:val="([^"]+)"/.exec(m[2])?.[1],
      level: hs ? Number(hs) : ol !== undefined ? Number(ol) + 1 : undefined,
    });
  }
  return map;
}

function readLevels(xml: string): (numId: string, ilvl: number) => { fmt: string; text: string } | null {
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
    return { fmt: /<w:numFmt w:val="(\w+)"/.exec(lvl)?.[1] ?? "decimal", text: /<w:lvlText w:val="([^"]*)"/.exec(lvl)?.[1] ?? "" };
  };
}

export async function materializeHeadingNumbers(
  zip: JSZip,
  nodes: (ParaNode | TableNode)[],
  chapterIds: Set<number>,
  stylesXml: string,
): Promise<{ count: number; regions: VerifyRegion[] }> {
  const out = { count: 0, regions: [] as VerifyRegion[] };
  const numXml = await zip.file("word/numbering.xml")?.async("string");
  if (!numXml) return out;
  const styles = readStyles(stylesXml);
  const lvlOf = readLevels(numXml);
  const fromStyle = <K extends keyof StyleInfo>(styleId: string | null, key: K): StyleInfo[K] | undefined => {
    let s = styleId;
    for (let d = 0; s && d < 8; d++) {
      const st = styles.get(s);
      if (!st) return undefined;
      if (st[key] !== undefined) return st[key];
      s = st.basedOn ?? null;
    }
    return undefined;
  };

  let chapter = 0;
  let sec = 0;
  let sub = 0;
  for (const n of nodes) {
    if (n.kind !== "paragraph") continue;
    const t = n.text.trim();
    if (!t) continue;
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(n.xml)?.[0] ?? "";
    const directOl = /<w:outlineLvl w:val="(\d)"/.exec(pPr)?.[1];
    const level = directOl !== undefined ? Number(directOl) + 1 : fromStyle(n.styleId, "level");

    // Bab: "BAB II …", atau romawi saja pada judul bab (Heading 1 / hasil unifyChapterHeadings).
    const cm = CHAPTER_RE.exec(t);
    const bm = !cm && (chapterIds.has(n.id) || level === 1) ? BARE_RE.exec(t) : null;
    if (cm || bm) {
      const v = cm ? (/^\d+$/.test(cm[1]) ? Number(cm[1]) : fromRoman(cm[1])) : fromRoman(bm![1]);
      if (v > 0 && v < 40) { chapter = v; sec = 0; sub = 0; }
      continue;
    }
    if (level !== 2 && level !== 3) continue;

    // Subbab bernomor ketikan: penghitung disamakan dengannya.
    const typed = TYPED_RE.exec(t);
    if (typed) {
      const nums = typed[1].split(".").map(Number);
      sec = nums[1] ?? sec;
      sub = nums.length >= 3 ? nums[2] : 0;
      continue;
    }

    // Penomoran otomatis Word (langsung di paragraf atau dari style-nya).
    const np = /<w:numPr>([\s\S]*?)<\/w:numPr>/.exec(pPr)?.[1];
    const numId = np ? /<w:numId w:val="(\d+)"/.exec(np)?.[1] : fromStyle(n.styleId, "numId");
    if (!numId || numId === "0" || !chapter) continue;
    const il = np ? /<w:ilvl w:val="(\d+)"/.exec(np)?.[1] : undefined;
    const ilvl = il !== undefined ? Number(il) : fromStyle(n.styleId, "ilvl") ?? 0;
    const lvl = lvlOf(numId, ilvl);
    if (!lvl || !/^decimal/.test(lvl.fmt) || !/%\d\.%\d/.test(lvl.text)) continue;

    let num: string;
    if (level === 2) { sec++; sub = 0; num = `${chapter}.${sec}`; }
    else { sub++; num = `${chapter}.${Math.max(1, sec)}.${sub}`; }

    const before = paragraphLine(n.xml);
    const rPr = /<w:r(?:\s[^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(n.xml)?.[1] ?? "";
    const startsTab = /^[ \u00a0]*\t/.test(leadingRaw(n.xml, 4));
    let xml = withProps(n.xml, "w:pPr", { "w:numPr": `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>` }, PPR_ORDER);
    const e = xml.indexOf("</w:pPr>");
    const at = e !== -1 ? e + 8 : xml.indexOf(">") + 1;
    xml = xml.slice(0, at) + `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(num)}</w:t>${startsTab ? "" : "<w:tab/>"}</w:r>` + xml.slice(at);
    n.xml = xml;
    n.hasNumbering = false;
    n.text = `${num}\t${t}`;
    out.regions.push({ beforeLines: [before].filter(Boolean), afterLines: [paragraphLine(n.xml)].filter(Boolean), beforeText: "", afterText: "" });
    out.count++;
  }
  return out;
}