/**
 * Subbab bernomor otomatis Word ("1.1" dari penomoran Heading 2, teks paragrafnya hanya
 * "Latar Belakang") → nomornya ditulis menjadi teks ("1.1⇥Latar Belakang") dan penomoran
 * otomatisnya dimatikan. Setelah itu subbab dikenali seperti subbab ketikan, nomornya ikut
 * diperiksa, dan Daftar Isi memuatnya dengan benar.
 *
 * Nomor yang ditulis = nomor yang ditampilkan Word di dokumen asli (numbering-sim). Hanya
 * penomoran desimal bertingkat ("%1.%2", "%1.%2.%3") yang ditulis ulang, dan per daftar Word
 * utuh-atau-tidak-sama-sekali. Dijalankan setelah unifyChapterHeadings dan sebelum outline.
 */
import type JSZip from "jszip";
import { escapeXml, withProps, PPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";
import { leadingRaw } from "./lists";
import { paragraphLine, type VerifyRegion } from "./verify";
import { createNumberingSim, type Displayed } from "./numbering-sim";

const HEAD_STYLE_RE = /^(?:heading|judul)\s*([1-6])$/i;
const TYPED_RE = /^\s*(\d+(?:\.\d+){1,4})[.)]?\s/;
const CHAPTER_RE = /^\s*(?:bab|chapter)\s+([ivxlcdm]+|\d+)\b/i;
/** Teks nomor bertingkat bertitik, bagian berupa placeholder (%1) atau angka tetap (1). */
const DOTTED_LVL_RE = /^(?:\d+|%\d)(?:\.(?:\d+|%\d))+\.?$/;

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
  /** Nomor yang DITAMPILKAN Word di dokumen asli (displayedLabels pada model sebelum diubah). */
  displayed: Map<ParaNode, Displayed>,
  /** Semua paragraf yang masih ada di model sekarang (termasuk di tabel). */
  present: ParaNode[],
  /** Paragraf lampiran yang dibekukan: tidak boleh diubah, jadi daftarnya juga tidak diubah. */
  frozen: Set<ParaNode> = new Set(),
): Promise<{ count: number; skippedLists: number; regions: VerifyRegion[]; written: Set<ParaNode> }> {
  const out = { count: 0, skippedLists: 0, regions: [] as VerifyRegion[], written: new Set<ParaNode>() };
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

  // Nomor yang ditulis = nomor yang DITAMPILKAN Word di dokumen asli (simulasi numbering.xml),
  // bukan hitungan ulang dari konteks bab: urutan XML tidak selalu sama dengan urutan tampil
  // (mis. hasil konversi PDF yang memakai kolom: "5.1 Kesimpulan" ada SEBELUM judul "BAB V").
  //
  // Satu daftar Word (abstractNum) dipakai bersama oleh beberapa tingkat (Heading 2 "2.1",
  // Heading 3 "2.1.4", Heading 4 "2.1.4.1") dan Word menghitung nomor tingkat bawah dari penghitung
  // tingkat atasnya. Jadi SATU daftar diubah UTUH (semua paragrafnya jadi nomor ketikan) atau TIDAK
  // SAMA SEKALI — mengubah sebagian membuat nomor paragraf lain di daftar itu berubah.
  void chapterIds;
  type Cand = { n: ParaNode; t: string; num: string; abs: string };
  const cands: Cand[] = [];
  const bad = new Set<string>();
  for (const n of nodes) {
    if (n.kind !== "paragraph") continue;
    const d = displayed.get(n);
    if (!d) continue;
    const t = n.text.trim();
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(n.xml)?.[0] ?? "";
    const directOl = /<w:outlineLvl w:val="(\d)"/.exec(pPr)?.[1];
    const level = directOl !== undefined ? Number(directOl) + 1 : fromStyle(n.styleId, "level");
    const lvl = lvlOf(d.numId, d.ilvl);
    // Penomoran desimal bertingkat bertitik: "%1.%2", "%1.%2.%3.%4", juga dengan awalan bab tetap ("2.%1").
    const dotted = Boolean(lvl) && /^decimal/.test(lvl!.fmt) && DOTTED_LVL_RE.test(lvl!.text.trim()) && /%\d/.test(lvl!.text);
    if (!dotted) continue;                       // daftar biasa (a., 1), •) — bukan urusan modul ini
    const label = d.label.trim().replace(/\.$/, "");
    const headingLike = t.length > 0 && t.length <= 150 && !/[.!?;:,]$/.test(t);
    const isHeading = (level !== undefined && level >= 2 && level <= 6) || headingLike;
    const ok = isHeading && !CHAPTER_RE.test(t) && !TYPED_RE.test(t) && /^\d+(?:\.\d+){1,4}$/.test(label);
    if (!ok || frozen.has(n)) { bad.add(d.abs); continue; }
    cands.push({ n, t, num: label, abs: d.abs });
  }
  // Daftar yang sama kadang juga dipakai paragraf lain (mis. butir "1)" di tingkat bawah). Dicoba:
  // nomor otomatis judul dimatikan → nomor yang tampil di SEMUA paragraf bernomor lain disimulasikan
  // ulang; daftar diterima hanya bila tidak ada satu pun nomor lain yang berubah.
  const simLabels = (skip: Set<ParaNode>) => {
    const sim = createNumberingSim(numXml, stylesXml);
    const m = new Map<ParaNode, string>();
    for (const p of present) {
      if (skip.has(p)) continue;
      const r = sim.next(p.xml, p.styleId);
      if (r) m.set(p, r.label);
    }
    return m;
  };
  const baseline = simLabels(new Set());
  let accepted = new Set<ParaNode>();
  for (const abs of new Set(cands.map((c) => c.abs))) {
    if (bad.has(abs)) continue;
    const trial = new Set([...accepted, ...cands.filter((c) => c.abs === abs).map((c) => c.n)]);
    const after = simLabels(trial);
    const changed = [...baseline].find(([p, l]) => !trial.has(p) && after.get(p) !== l);
    if (changed) {
     
      bad.add(abs);
    } else accepted = trial;
  }
  out.skippedLists = new Set(cands.filter((c) => bad.has(c.abs)).map((c) => c.abs)).size;
  for (const { n, t, num, abs } of cands) {
    if (bad.has(abs)) continue;
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
    out.written.add(n);
  }
  return out;
}