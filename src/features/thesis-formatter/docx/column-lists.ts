/**
 * Daftar berkolom (mis. pedoman UNPAD Fakultas Peternakan):
 *
 *   DAFTAR ISI                              DAFTAR TABEL / ILUSTRASI / LAMPIRAN
 *   Bab                          Halaman    Nomor                        Halaman
 *         KATA PENGANTAR ........... i        1     Judul tabel ............ 8
 *   I     PENDAHULUAN                         2     Judul panjang yang
 *         1.1 Latar Belakang ....... 1              berlanjut .............. 22
 *
 * Field TOC Word tidak bisa membuat kolom Bab terpisah, menghapus nomor halaman HANYA pada
 * baris bab, atau membuang label "Tabel" — jadi daftar disusun sebagai paragraf biasa.
 * Nomor halaman tetap otomatis: field PAGEREF ke bookmark di setiap judul/keterangan
 * (diperbarui saat dokumen dibuka / F9, dan bisa diklik). Nomor halaman lama dipakai
 * sebagai isi tampilan awal.
 *
 * Teks entri diambil dari ISI paragraf judulnya (bukan disusun ulang): huruf persis sama
 * dan cetak miring sebagian (nama ilmiah, istilah asing) ikut terbawa.
 */
import { escapeXml } from "./xml";
import { parseBlocks, serializeDocument, type DocModel, type ParaNode, type TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { paragraphLine, type VerifyRegion } from "./verify";
import { unwrapTocControls, type TocReport } from "./toc";

type ListKind = "contents" | "tables" | "figures" | "appendices";
/** Potongan teks dengan tanda miring. */
type Seg = { t: string; i: boolean };
type Row =
  | { type: "front"; text: string; target: ParaNode }
  | { type: "chapter"; label: string; text: string; segs: Seg[] }
  | { type: "section" | "subsection"; num: string; text: string; segs: Seg[]; target: ParaNode }
  | { type: "item"; num: string; text: string; segs: Seg[]; target: ParaNode };

/** Halaman awal yang tidak dicantumkan di Daftar Isi. */
const FRONT_SKIP_RE = /\b(halaman\s+(judul|sampul)|sampul|persetujuan|pengesahan|pernyataan|motto|moto|persembahan|berita\s+acara)\b/i;
const TOC_INSTR_RE = /<w:instrText[^>]*>\s*TOC\b|<w:fldSimple\b[^>]*w:instr="\s*TOC\b/;
const CH_RE = /^\s*(?:bab\s+)?([IVXLC]{1,6}|\d{1,2})\b[.:]?\s+(\S.*)$/i;
const SEC_RE = /^\s*(\d+(?:\.\d+){1,3})[.)]?\s+(\S.*)$/;
const CAP_RE = /^\s*[A-Za-z]+\s*\.?\s*((?:[A-Za-z]{1,2}[.\-]?)?\d+(?:[.\-]\d+)*)\s*[.:\u2013\u2014-]?\s*(\S.*)$/;
const APPX_RE = /^\s*lampiran\s+(\d{1,2}|[A-Za-z]|[IVXLCivxlc]{1,5})(?![A-Za-z0-9])\s*[.:\-\u2013\u2014]?\s*(\S.*)$/i;

/** Awalan yang dibuang dari teks paragraf asli (nomor bab/subbab/keterangan/lampiran). */
const CH_PREFIX = /^\s*(?:bab\s+)?(?:[IVXLC]{1,6}|\d{1,2})\b[.:]?\s+/i;
const SEC_PREFIX = /^\s*\d+(?:\.\d+){1,3}[.)]?\s+/;
const CAP_PREFIX = /^\s*[A-Za-z]+\s*\.?\s*(?:[A-Za-z]{1,2}[.\-]?)?\d+(?:[.\-]\d+)*\s*[.:\u2013\u2014-]?\s*/;
const APPX_PREFIX = /^\s*lampiran\s+(?:\d{1,2}|[A-Za-z]|[IVXLCivxlc]{1,5})(?![A-Za-z0-9])\s*[.:\-\u2013\u2014]?\s*/i;

const decode = (s: string) => s
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

/** Teks paragraf per potongan miring/tegak (dari w:t; tab & ganti baris dibaca spasi). */
function segmentsOf(xml: string): Seg[] {
  const out: Seg[] = [];
  const body = xml.slice(Math.max(0, xml.indexOf("</w:pPr>")));
  for (const m of body.matchAll(/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g)) {
    const rPr = /^\s*<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(m[1])?.[1] ?? "";
    const it = /<w:i(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(rPr);
    let t = "";
    for (const x of m[1].matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>|<w:br\b[^>]*\/>/g)) {
      t += x[1] !== undefined ? decode(x[1]) : " ";
    }
    if (!t) continue;
    const last = out[out.length - 1];
    if (last && last.i === it) last.t += t;
    else out.push({ t, i: it });
  }
  return out;
}

/** Potongan setelah awalan `re`; spasi dirapikan; miring seluruhnya = tegak (gaya, bukan istilah). */
function afterPrefix(segs: Seg[], re: RegExp | null): Seg[] {
  const full = segs.map((s) => s.t).join("");
  let cut = re ? re.exec(full)?.[0].length ?? 0 : 0;
  let out: Seg[] = [];
  for (const s of segs) {
    if (cut >= s.t.length) { cut -= s.t.length; continue; }
    out.push({ t: s.t.slice(cut).replace(/\s+/g, " "), i: s.i });
    cut = 0;
  }
  if (out.length) {
    out[0].t = out[0].t.replace(/^\s+/, "");
    out[out.length - 1].t = out[out.length - 1].t.replace(/\s+$/, "");
  }
  out = out.filter((s) => s.t);
  if (out.length && out.every((s) => s.i)) out = out.map((s) => ({ ...s, i: false }));
  return out;
}

const joined = (segs: Seg[]) => segs.map((s) => s.t).join("");

/** Lebar glyph (per 1000 em) untuk mengukur kata kepala kolom ("Bab", "Nomor", "Halaman"); huruf lain ±0,52 em. */
const GLYPHS: Record<string, { regular: Record<string, number>; bold: Record<string, number> }> = {
  "Times New Roman": {
    regular: { H: 722, a: 444, l: 278, m: 778, n: 500, B: 667, b: 500, N: 722, o: 500, r: 333 },
    bold: { H: 778, a: 500, l: 278, m: 833, n: 556, B: 667, b: 556, N: 722, o: 500, r: 444 },
  },
  Arial: {
    regular: { H: 722, a: 556, l: 222, m: 833, n: 556, B: 667, b: 556, N: 722, o: 556, r: 333 },
    bold: { H: 722, a: 556, l: 278, m: 889, n: 611, B: 722, b: 611, N: 722, o: 611, r: 389 },
  },
  Calibri: {
    regular: { H: 623, a: 479, l: 229, m: 799, n: 525, B: 544, b: 525, N: 646, o: 527, r: 349 },
    bold: { H: 631, a: 494, l: 246, m: 813, n: 537, B: 561, b: 537, N: 659, o: 538, r: 355 },
  },
};

function textTwips(text: string, font: string, halfPts: number, bold: boolean): number {
  const table = GLYPHS[font] ?? GLYPHS["Times New Roman"];
  const w = bold ? table.bold : table.regular;
  let units = 0;
  for (const ch of text) units += w[ch] ?? 520;
  return Math.ceil((units / 1000) * halfPts * 10);
}

/** Lebar nomor ("2.1.10") dalam twip: angka & titik punya lebar baku per font. */
const NUM_GLYPH: Record<string, { d: number; dot: number }> = {
  "Times New Roman": { d: 500, dot: 250 },
  Arial: { d: 556, dot: 278 },
  Calibri: { d: 507, dot: 252 },
};
function numTwips(num: string, font: string, halfPts: number): number {
  const g = NUM_GLYPH[font] ?? NUM_GLYPH["Times New Roman"];
  let units = 0;
  for (const ch of num) units += /\d/.test(ch) ? g.d : ch === "." ? g.dot : 500;
  return Math.ceil((units / 1000) * halfPts * 10);
}

function fromRoman(s: string): number {
  const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100 };
  const a = s.toLowerCase();
  let t = 0;
  for (let k = 0; k < a.length; k++) { const c = v[a[k]] ?? 0; const n = v[a[k + 1]] ?? 0; t += c < n ? -c : c; }
  return t;
}
function toRoman(n: number): string {
  const map: [number, string][] = [[10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let out = "";
  for (const [v, s] of map) while (n >= v) { out += s; n -= v; }
  return out;
}

function kindOf(t: string): ListKind | null {
  if (/\bisi\b/i.test(t)) return "contents";
  if (/tabel/i.test(t)) return "tables";
  if (/gambar|ilustrasi|grafik|bagan/i.test(t)) return "figures";
  if (/lampiran/i.test(t)) return "appendices";
  return null;
}

/** Kunci pencocokan: tanpa label/nomor/romawi di depan, huruf kecil, tanda baca dibuang. */
const keyOf = (t: string) => t
  .replace(/^[\s\t]*(?:(?:bab|tabel|table|gambar|ilustrasi|figure|lampiran)\s+)?(?:\d+(?:\.\d+)*\.?|[ivxlc]{1,6}\.?)?\s+/i, "")
  .toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

/** Nomor halaman dari daftar lama (entri panjang yang terbelah dua baris digabung dulu). */
function oldPages(paras: ParaNode[]): { key: string; page: string }[] {
  const out: { key: string; page: string }[] = [];
  let pending = "";
  for (const p of paras) {
    const t = p.text.replace(/\s+$/, "");
    if (!t.trim()) continue;
    const m = /^(.*?)(?:\t|\.{2,}|\u2026|\s{2,})+\s*([0-9]+|[ivxlcdm]+)$/i.exec(t);
    if (m) { out.push({ key: keyOf(`${pending} ${m[1]}`.trim()), page: m[2] }); pending = ""; }
    else pending = `${pending} ${t}`.trim();
  }
  return out;
}

export function buildColumnLists(
  model: DocModel,
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  o: { width: number; font: string; halfPts: number; bodyLine: number; upper: boolean; exact: boolean },
): { report: TocReport; regions: VerifyRegion[] } {
  const report: TocReport = { tocs: 0, entries: 0, replaced: 0, unwrapped: unwrapTocControls(model) };
  const regions: VerifyRegion[] = [];

  const lists: { head: ParaNode; kind: ListKind; idx: number }[] = [];
  nodes.forEach((n, i) => {
    if (n.kind !== "paragraph" || outline.items[i]?.role !== "list_heading") return;
    const k = kindOf(n.text);
    if (k) lists.push({ head: n, kind: k, idx: i });
  });
  if (!lists.length) return { report, regions };

  /* ---- bookmark di setiap judul / keterangan ---- */
  let bm = 1 + [...serializeDocument(model).matchAll(/<w:bookmarkStart\b[^>]*w:id="(\d+)"/g)]
    .reduce((m, x) => Math.max(m, Number(x[1])), 0);
  const marks = new Map<ParaNode, string>();
  const mark = (p: ParaNode) => {
    const had = marks.get(p);
    if (had) return had;
    const id = bm++;
    const name = `_FmtRef${id}`;
    const e = p.xml.indexOf("</w:pPr>");
    const at = e !== -1 ? e + 8 : p.xml.indexOf(">") + 1;
    p.xml = p.xml.slice(0, at) + `<w:bookmarkStart w:id="${id}" w:name="${name}"/>` +
      p.xml.slice(at).replace(/<\/w:p>$/, `<w:bookmarkEnd w:id="${id}"/></w:p>`);
    marks.set(p, name);
    return name;
  };

  /* ---- entri ---- */
  const contents: Row[] = [];
  const tables: Row[] = [];
  const figures: Row[] = [];
  const appendices: Row[] = [];
  let appendixListed = false;
  const listAppendix = (n: ParaNode) => {
    if (appendixListed) return;
    appendixListed = true;
    contents.push({ type: "front", text: "LAMPIRAN", target: n });
  };
  const upperSegs = (segs: Seg[]) => (o.upper ? segs.map((s) => ({ ...s, t: s.t.toUpperCase() })) : segs);
  /** Baris bab. `segs`: teks judul dari paragrafnya; `ownPrefix`: segs masih memuat nomor bab. */
  const chapterRow = (t: string, segs: Seg[], ownPrefix: boolean): Row => {
    const m = CH_RE.exec(t);
    const v = m ? (/^\d+$/.test(m[1]) ? Number(m[1]) : fromRoman(m[1])) : 0;
    const valid = Boolean(m && v > 0 && v < 40 && (/^\d+$/.test(m[1]) || toRoman(v) === m[1].toUpperCase()));
    const title = upperSegs(afterPrefix(segs, valid && ownPrefix ? CH_PREFIX : null));
    return {
      type: "chapter",
      label: valid ? (/^\d+$/.test(m![1]) ? toRoman(v) : m![1].toUpperCase()) : "",
      text: joined(title) || (valid ? m![2] : t),
      segs: title,
    };
  };

  nodes.forEach((n, i) => {
    if (n.kind !== "paragraph") return;
    const it = outline.items[i];
    if (!it || it.zone === "cover") return;
    const t = n.text.replace(/\s+/g, " ").trim();
    if (!t) return;
    switch (it.role) {
      case "front_heading":
        if (!FRONT_SKIP_RE.test(t)) contents.push({ type: "front", text: t.toUpperCase(), target: n });
        return;
      case "list_heading":
      case "reference_heading":
        contents.push({ type: "front", text: t.toUpperCase(), target: n });
        return;
      case "appendix_heading": {
        listAppendix(n);
        const m = APPX_RE.exec(t);
        if (m && /^[A-Z0-9(\u201c"]/.test(m[2])) {
          const segs = afterPrefix(segmentsOf(n.xml), APPX_PREFIX);
          appendices.push({ type: "item", num: m[1], text: joined(segs) || m[2], segs, target: n });
        }
        return;
      }
      case "chapter_label": {
        if (!it.pairedTitleNext) return;
        let j = i + 1;
        while (j < nodes.length && outline.items[j]?.role === "empty") j++;
        const q = nodes[j];
        const title = q && q.kind === "paragraph" ? q.text.replace(/\s+/g, " ").trim() : "";
        const segs = q && q.kind === "paragraph" ? segmentsOf(q.xml) : [];
        contents.push(chapterRow(`${t.replace(/^\s*bab\s+/i, "")} ${title}`, segs, false));
        return;
      }
      case "chapter":
        if (it.continuesChapterLabel) return;
        contents.push(chapterRow(t, segmentsOf(n.xml), true));
        return;
      case "section":
      case "subsection": {
        if (it.zone !== "main") return;
        const m = SEC_RE.exec(n.text.replace(/\t/g, " "));
        const segs = afterPrefix(segmentsOf(n.xml), m ? SEC_PREFIX : null);
        contents.push({ type: it.role, num: m?.[1] ?? "", text: joined(segs) || (m ? m[2] : t).trim(), segs, target: n });
        return;
      }
      case "table_caption":
      case "figure_caption": {
        if (it.zone !== "main") return;
        const m = CAP_RE.exec(t);
        if (!m) return;
        const segs = afterPrefix(segmentsOf(n.xml), CAP_PREFIX);
        (it.role === "table_caption" ? tables : figures).push({ type: "item", num: m[1], text: joined(segs) || m[2], segs, target: n });
        return;
      }
      default:
        // "Lampiran 1. Judul" yang bukan judul halaman (judul panjang di atas isi lampiran).
        if (it.zone === "back" && it.role === "body" && t.length <= 250) {
          const m = APPX_RE.exec(t);
          if (m && /^[A-Z0-9(\u201c"]/.test(m[2]) && !/[.!?]$/.test(m[2])) {
            listAppendix(n);
            const segs = afterPrefix(segmentsOf(n.xml), APPX_PREFIX);
            appendices.push({ type: "item", num: m[1], text: joined(segs) || m[2], segs, target: n });
          }
        }
    }
  });

  /* ---- penyusun paragraf ---- */
  const W = o.width;
  const COL = 709;   // kolom nomor bab / nomor tabel (1,25 cm)
  const SEC = 567;   // lebar nomor subbab "1.1"
  const SUB = 850;   // lebar nomor anak subbab "2.1.1"
  const half = Math.round(o.bodyLine / 2);
  const fn = escapeXml(o.font);
  const rPr = (bold: boolean, italic = false) =>
    `<w:rPr><w:rFonts w:ascii="${fn}" w:hAnsi="${fn}" w:cs="${fn}" w:eastAsia="${fn}"/>${bold ? "<w:b/><w:bCs/>" : ""}` +
    `${italic ? "<w:i/><w:iCs/>" : ""}<w:sz w:val="${o.halfPts}"/><w:szCs w:val="${o.halfPts}"/></w:rPr>`;
  const run = (text: string, bold = false, italic = false) =>
    `<w:r>${rPr(bold, italic)}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
  /** Teks entri: potongan miring dari judul aslinya dipertahankan. */
  const runs = (segs: Seg[] | undefined, text: string, bold = false) =>
    segs && segs.length ? segs.map((s) => run(s.t, bold, s.i)).join("") : run(text, bold);
  const tab = (bold = false) => `<w:r>${rPr(bold)}<w:tab/></w:r>`;
  const ref = (name: string, cached: string, bold = false) =>
    `<w:r>${rPr(bold)}<w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>` +
    `<w:r>${rPr(bold)}<w:instrText xml:space="preserve"> PAGEREF ${name} \\h </w:instrText></w:r>` +
    `<w:r>${rPr(bold)}<w:fldChar w:fldCharType="separate"/></w:r>` +
    (cached ? run(cached, bold) : "") +
    `<w:r>${rPr(bold)}<w:fldChar w:fldCharType="end"/></w:r>`;
  const lineAttr = (line: number) => o.exact
    ? `w:line="${Math.round((line / 240) * o.halfPts * 10)}" w:lineRule="exact"`
    : `w:line="${line}" w:lineRule="auto"`;
  /**
 * Kolom "Halaman" selebar kata "Halaman" di baris kepala (rata kanan di margin):
 *   judul ........ (titik-titik berhenti tepat sebelum huruf "H")       12
 *                                                     nomor rata kanan ↑ = huruf "n"
 * Teks judul (termasuk baris lanjutan) tidak pernah masuk ke kolom ini (indentasi kanan).
 */
  const PAGE_LABEL = "Halaman";
  const GAP = 57; // ±1 mm: titik terakhir sedikit sebelum huruf "H"
  const dotEndFor = (bold: boolean) => W - textTwips(PAGE_LABEL, o.font, o.halfPts, bold) - GAP;
  const para = (
    x: { left: number; hanging?: number; before?: number; line?: number; keepNext?: boolean; tabs?: string; dotEnd?: number; noLeader?: boolean; center?: number },
    inner: string,
  ) => {
    const column = x.dotEnd !== undefined && !x.noLeader
      ? `<w:tab w:val="left" w:leader="dot" w:pos="${x.dotEnd}"/><w:tab w:val="right" w:pos="${W}"/>`
      : "";
    const tabs = x.tabs ??
      `${x.center ? `<w:tab w:val="center" w:pos="${x.center}"/>` : ""}${x.hanging ? `<w:tab w:val="left" w:pos="${x.left}"/>` : ""}${column}`;
    const right = x.dotEnd !== undefined ? Math.max(0, W - x.dotEnd) : 0;
    return `<w:p><w:pPr>${x.keepNext ? "<w:keepNext/>" : ""}${tabs ? `<w:tabs>${tabs}</w:tabs>` : ""}` +
      `<w:spacing w:before="${x.before ?? 0}" w:after="0" ${lineAttr(x.line ?? 240)}/>` +
      `<w:ind w:left="${x.left}" w:right="${right}" ${x.hanging ? `w:hanging="${x.hanging}"` : `w:firstLine="0"`}/>` +
      `<w:jc w:val="left"/></w:pPr>${inner}</w:p>`;
  };
  const header = (l: string, r: string, bold: boolean) =>
    para({ left: 0, keepNext: true, tabs: `<w:tab w:val="right" w:pos="${W}"/>` }, run(l, bold) + tab(bold) + run(r, bold));

  const pageOf = (cache: { key: string; page: string }[], text: string) => {
    const k = keyOf(text);
    if (!k) return "";
    const hit = cache.find((c) => c.key === k) ??
      cache.find((c) => c.key.length >= 8 && k.length >= 8 && (c.key.startsWith(k) || k.startsWith(c.key)));
    return hit?.page ?? "";
  };

  const contentsXml = (rows: Row[], cache: { key: string; page: string }[]) => {
    const dotEnd = dotEndFor(true);
    // Angka Romawi bab di tengah kolom, tepat di bawah tengah kata "Bab".
    const numCenter = Math.round(textTwips("Bab", o.font, o.halfPts, true) / 2);
    // Jarak nomor ke judul secukupnya ("1.1  Latar Belakang"): kolom judul tiap tingkat = nomor
    // terlebar di tingkat itu + ±2 spasi, jadi judul setingkat tetap lurus. Nomor anak subbab
    // mulai sejajar huruf pertama judul subbab.
    const NUM_GAP = 113;
    const widest = (t: "section" | "subsection") =>
      Math.max(0, ...rows.map((r) => (r.type === t && "num" in r && r.num ? numTwips(r.num, o.font, o.halfPts) : 0)));
    const secW = widest("section");
    const subW = widest("subsection");
    const secText = COL + (secW ? secW + NUM_GAP : SEC);
    const subText = secText + (subW ? subW + NUM_GAP : SUB);
    const out = [header("Bab", PAGE_LABEL, true)];
    let prev: Row["type"] | null = null;
    for (const r of rows) {
      if (r.type === "front") {
        const before = prev === null ? half : prev !== "front" ? half : 0;
        out.push(para({ left: COL, before, line: o.bodyLine, dotEnd },
          run(r.text, true) + tab(true) + tab(true) + ref(mark(r.target), pageOf(cache, r.text), true)));
      } else if (r.type === "chapter") {
        // Baris bab: tanpa titik-titik & tanpa nomor halaman; tetap tidak masuk kolom "Halaman".
        out.push(para({ left: COL, hanging: COL, before: half, line: o.bodyLine, keepNext: true, noLeader: true, dotEnd, center: numCenter },
          tab(true) + (r.label ? run(r.label, true) : "") + tab(true) + runs(r.segs, r.text, true)));
      } else if (r.type === "section" || r.type === "subsection") {
        const sub = r.type === "subsection";
        const left = sub ? subText : secText;
        out.push(para({ left, hanging: sub ? subText - secText : secText - COL, dotEnd },
          (r.num ? run(r.num) : "") + tab() + runs(r.segs, r.text) + tab() + tab() +
          ref(mark(r.target), pageOf(cache, `${r.num} ${r.text}`))));
      }
      prev = r.type;
    }
    return out.join("");
  };

  const itemsXml = (rows: Row[], cache: { key: string; page: string }[]) => {
    const dotEnd = dotEndFor(false);
    // Nomor di tengah kolom, tepat di bawah tengah kata "Nomor".
    const numCenter = Math.round(textTwips("Nomor", o.font, o.halfPts, false) / 2);
    const out = [header("Nomor", PAGE_LABEL, false)];
    rows.forEach((r, k) => {
      if (r.type !== "item") return;
      out.push(para({ left: COL, hanging: COL, before: k === 0 ? 240 : 120, dotEnd, center: numCenter },
        tab() + run(r.num) + tab() + runs(r.segs, r.text) + tab() + tab() + ref(mark(r.target), pageOf(cache, r.text))));
    });
    return out.join("");
  };

  /* ---- ganti isi tiap halaman daftar ---- */
  let nextId = 5_000_000 + nodes.reduce((m, n) => Math.max(m, n.id), 0);
  for (const l of lists) {
    const region: ParaNode[] = [];
    for (let j = l.idx + 1; j < nodes.length; j++) {
      const m = nodes[j];
      const r = outline.items[j]?.role;
      if (m.kind !== "paragraph") break;
      if (r === "list_entry" || r === "empty" || TOC_INSTR_RE.test(m.xml)) region.push(m);
      else break;
    }
    const rows = l.kind === "contents" ? contents : l.kind === "tables" ? tables : l.kind === "figures" ? figures : appendices;
    const cache = oldPages(region);
    let xml = l.kind === "contents" ? contentsXml(rows, cache) : itemsXml(rows, cache);

    // Section break yang menempel di entri lama dipindah ke baris daftar terakhir.
    const sect = [...region].reverse().map((p) => /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(p.xml)?.[0]).find(Boolean);
    if (sect) {
      const at = xml.lastIndexOf("</w:pPr>");
      xml = xml.slice(0, at) + sect + xml.slice(at);
    }

    if (model.body.indexOf(l.head) === -1) continue;
    const fresh = parseBlocks(xml, nextId) as ParaNode[];
    nextId += fresh.length + 10;
    const drop = new Set<ParaNode>(region);
    model.body = model.body.filter((b) => !(b.kind === "paragraph" && drop.has(b)));
    model.body.splice(model.body.indexOf(l.head) + 1, 0, ...fresh);

    regions.push({
      beforeLines: region.map((p) => paragraphLine(p.xml)).filter(Boolean),
      afterLines: fresh.map((p) => paragraphLine(p.xml)).filter(Boolean),
      beforeText: "",
      afterText: "",
    });
    report.tocs++;
    report.entries += rows.length;
  }
  return { report, regions };
}