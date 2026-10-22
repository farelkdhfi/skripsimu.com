/**
 * Segala hal seputar tabel:
 *  - tata letak tabel (lebar, border, header, perataan sel)
 *  - keterangan (judul) tabel: rata kiri di atas tabel, Title Case, tidak tebal
 *  - catatan di bawah tabel (Sumber/Keterangan/Catatan): 10 pt
 */
import { escapeXml, mergeProps, withProps, PPR_ORDER, TBLPR_ORDER, TCPR_ORDER, TRPR_ORDER } from "./xml";
import { cellParagraphs, type CellNode, type DocModel, type ParaNode, type TableNode } from "./document-model";
import { formatParagraph, type ApplyReport } from "./apply-format";
import { paragraphLine, type VerifyRegion } from "./verify";
import type { FormatRules, ParaFormat, RunFormat, TableFormat } from "../rules/format-rules";
import type { Outline, OutlineItem } from "../structure/outline";
import { leadingRaw, removeLeading } from "./lists";

/* ------------------------------------------------------------------ */
/* Tata letak tabel                                                     */
/* ------------------------------------------------------------------ */

const NUMERIC_RE = /^[\s\d.,%()+\-\u2013\u2014/]+$/;

function isNumericColumn(rows: string[][], col: number): boolean {
  const vals = rows.slice(1).map((r) => (r[col] ?? "").trim()).filter(Boolean);
  if (vals.length < 2) return false;
  return vals.filter((v) => NUMERIC_RE.test(v) && /\d/.test(v)).length / vals.length >= 0.6;
}

function dxaOf(tcPr: string | null): number | null {
  if (!tcPr) return null;
  const m = /<w:tcW\s+[^>]*w:w="(\d+)"[^>]*w:type="dxa"/.exec(tcPr) ?? /<w:tcW\s+[^>]*w:type="dxa"[^>]*w:w="(\d+)"/.exec(tcPr);
  return m ? Number(m[1]) : null;
}

function formatCellParagraphs(cell: CellNode, align: ParaFormat["align"], run: RunFormat, t: TableFormat, rep: ApplyReport) {
  for (const p of cellParagraphs(cell)) {
    formatParagraph(
      p,
      { align, firstLine: 0, left: 0, line: t.lineSpacing, before: 0, after: 0, styleId: null },
      run,
      false,
      rep,
    );
  }
}

export type FormatTableCtx = {
  fmt: TableFormat;
  rules: FormatRules;
  usableWidth: number;
  texts: string[][];
  rep: ApplyReport;
};

export function normalizeTable(table: TableNode, t: FormatTableCtx, rep: ApplyReport) {
  const { fmt, rules, usableWidth, texts } = t;

  const colCount = Math.max(...table.rows.map((r) => r.cells.length), 0);
  const numeric = Array.from({ length: colCount }, (_, j) => isNumericColumn(texts, j));

  // Baris pertama dianggap header hanya bila isinya label, bukan angka.
  const hasHeaderRow =
    table.rows.length >= 2 &&
    (texts[0] ?? []).length > 0 &&
    (texts[0] ?? []).every((c) => !(NUMERIC_RE.test(c) && /\d/.test(c)));

  if (fmt.widthPct !== null) {
    // Urutan anak w:tblBorders wajib: top, left, bottom, right, insideH, insideV.
    const edge = (k: string, on: boolean) => on
      ? `<w:${k} w:val="single" w:sz="${fmt.borderSize}" w:space="0" w:color="000000"/>`
      : `<w:${k} w:val="nil"/>`;
    const borders = fmt.borders === "keep"
      ? undefined
      : `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"]
        .map((k) => edge(k, fmt.borders === "all" || k === "top" || k === "bottom"))
        .join("")}</w:tblBorders>`;

    const patched = mergeProps(table.props, {
      "w:tblW": `<w:tblW w:w="${Math.round(fmt.widthPct * 50)}" w:type="pct"/>`,
      "w:tblLayout": `<w:tblLayout w:type="autofit"/>`,
      "w:tblInd": `<w:tblInd w:w="0" w:type="dxa"/>`,
      "w:tblCellMar": `<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="${fmt.cellMargin}" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="${fmt.cellMargin}" w:type="dxa"/></w:tblCellMar>`,
      ...(borders === undefined ? {} : { "w:tblBorders": borders }),
    }, TBLPR_ORDER);
    table.props = `<w:tblPr>${patched.join("")}</w:tblPr>`;
  }

  // Skala w:tblGrid agar total lebarnya pas di dalam margin.
  if (table.grid && usableWidth > 0) {
    const widths = [...table.grid.matchAll(/<w:gridCol\s+w:w="(\d+)"\s*\/>/g)].map((m) => Number(m[1]));
    const total = widths.reduce((a, b) => a + b, 0);
    if (total > 0 && Math.abs(total - usableWidth) / usableWidth > 0.02) {
      const k = usableWidth / total;
      let idx = 0;
      table.grid = table.grid.replace(/<w:gridCol\s+w:w="\d+"\s*\/>/g, () =>
        `<w:gridCol w:w="${Math.max(200, Math.round(widths[idx++] * k))}"/>`);
    }
  }

  // Lebar kolom diturunkan dari w:tblGrid agar baris ber-gridSpan tetap sejajar.
  const gridW = table.grid
    ? [...table.grid.matchAll(/<w:gridCol\s+w:w="(\d+)"\s*\/>/g)].map((m) => Number(m[1]))
    : [];
  const gridTotal = gridW.reduce((a, b) => a + b, 0);

  table.rows.forEach((row, ri) => {
    const isHeader = ri === 0 && hasHeaderRow;
    const rowDxa = row.cells.map((c) => dxaOf(c.props));
    const rowTotal = rowDxa.reduce<number>((a, b) => a + (b ?? 0), 0);

    let col = 0;
    const spanOf = (c: CellNode) => {
      const m = c.props ? /<w:gridSpan\s+w:val="(\d+)"/.exec(c.props) : null;
      return m ? Number(m[1]) : 1;
    };
    const gridPct = row.cells.map((c) => {
      const span = spanOf(c);
      const w = gridW.slice(col, col + span).reduce((a, b) => a + b, 0);
      col += span;
      return gridTotal > 0 && w > 0 ? Math.round(5000 * w / gridTotal) : null;
    });

    if (isHeader && (fmt.repeatHeaderRow || fmt.boldHeaderRow)) {
      const patched = mergeProps(row.props, {
        ...(fmt.repeatHeaderRow ? { "w:tblHeader": "<w:tblHeader/>" } : {}),
        "w:cantSplit": "<w:cantSplit/>",
      }, TRPR_ORDER);
      row.props = `<w:trPr>${patched.join("")}</w:trPr>`;
    }

    row.cells.forEach((cell, ci) => {
      const pct = gridPct[ci]
        ?? (rowTotal > 0 && rowDxa[ci]
          ? Math.round(5000 * (rowDxa[ci] as number) / rowTotal)
          : Math.round(5000 / Math.max(1, row.cells.length)));

      const patched = mergeProps(cell.props, {
        ...(fmt.widthPct !== null ? { "w:tcW": `<w:tcW w:w="${pct}" w:type="pct"/>` } : {}),
        ...(fmt.clearShading ? { "w:shd": null } : {}),
        ...(fmt.borders !== "keep" ? {
          "w:tcBorders": ri === 0 && fmt.borders === "minimal" && table.rows.length > 1
            ? `<w:tcBorders><w:bottom w:val="single" w:sz="${fmt.borderSize}" w:space="0" w:color="000000"/></w:tcBorders>`
            : null,
        } : {}),
        "w:vAlign": `<w:vAlign w:val="center"/>`,
      }, TCPR_ORDER);
      cell.props = `<w:tcPr>${patched.join("")}</w:tcPr>`;

      const align = isHeader ? fmt.headerAlign : numeric[ci] ? fmt.numericAlign : "left";
      const noHint = { align: null, tabCount: 0, leadingTabsOnly: false, listMarker: false, allBold: false, length: 0, frontKind: null };
      const run = { ...rules.run("body", "main", noHint), size: fmt.fontSize, ...(isHeader && fmt.boldHeaderRow ? { bold: true } : {}) };
      formatCellParagraphs(cell, align, run, fmt, t.rep);

      for (const child of cell.children) {
        if (child.kind !== "table") continue;
        const childTexts = child.rows.map((r) => r.cells.map((cc) => cellParagraphs(cc).map((p) => p.text).join(" ").trim()));
        normalizeTable(child, { ...t, texts: childTexts }, rep);
      }
    });
  });

  rep.tablesNormalized++;
}

/* ------------------------------------------------------------------ */
/* Keterangan tabel: Title Case                                         */
/* ------------------------------------------------------------------ */

const SMALL_WORDS = new Set([
  "dan", "atau", "di", "ke", "dari", "yang", "untuk", "pada", "dengan", "terhadap", "dalam",
  "oleh", "serta", "sebagai", "per", "tentang", "antara", "bagi",
  "of", "and", "or", "the", "in", "on", "for", "to", "vs", "a", "an", "at", "by",
]);
const WORD_RE = /[A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u00FF][A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u00FF'\u2019]*/g;
const LABEL_RE = /^(tabel|table)\s*\.?\s*(?:[A-Za-z]{1,2}[.\-]?)?\d+(?:[.\-]\d+)*\s*[.:]?\s*/i;

/** "TABEL 3.5 KRITERIA RATA-RATA SKOR" → "Tabel 3.5 Kriteria Rata-Rata Skor". */
/** Title Case umum: kata sambung kecil, akronim dipertahankan kecuali teks kapital semua. */
export function titleCaseText(s: string, allCapsHint?: boolean): string {
  const letters = s.replace(/[^A-Za-z]/g, "");
  const allCaps = allCapsHint ?? (letters.length > 3 && letters === letters.toUpperCase());
  let first = true;
  return s.replace(WORD_RE, (w) => {
    const isFirst = first;
    first = false;
    if (!allCaps && w.length > 1 && (w === w.toUpperCase() || /[A-Z]/.test(w.slice(1)))) return w;
    const lower = w.toLowerCase();
    if (!isFirst && SMALL_WORDS.has(lower)) return lower;
    return lower[0].toUpperCase() + lower.slice(1);
  });
}

/** "TABEL 3.5 KRITERIA RATA-RATA SKOR" → "Tabel 3.5 Kriteria Rata-Rata Skor". */
export function titleCaseCaption(s: string): string {
  const letters = s.replace(/[^A-Za-z]/g, "");
  const allCaps = letters.length > 3 && letters === letters.toUpperCase();
  const label = LABEL_RE.exec(s);
  const head = label ? label[0].replace(/^(tabel|table)/i, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase()) : "";
  return head + titleCaseText(s.slice(label ? label[0].length : 0), allCaps);
}

const decode = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

/** Terapkan transformasi huruf ke teks paragraf tanpa merusak run/format. */
export function retext(xml: string, fn: (s: string) => string): string {
  const re = /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  const texts = [...xml.matchAll(re)].map((m) => decode(m[2]));
  const full = texts.join("");
  const out = fn(full);
  if (out === full || out.length !== full.length) return xml;
  let pos = 0;
  let k = 0;
  return xml.replace(re, (_m, attr) => {
    const t = texts[k++];
    const piece = out.slice(pos, pos + t.length);
    pos += t.length;
    return `<w:t${attr ?? ""}>${escapeXml(piece)}</w:t>`;
  });
}

/* ------------------------------------------------------------------ */
/* Keterangan di atas tabel + catatan di bawah tabel                    */
/* ------------------------------------------------------------------ */

const NOTE_HEAD_RE = /^(sumber|source|keterangan|ket\.|catatan|note|notes|data\s+diolah|diolah)\b/i;
const NOTE_SYMBOL_RE = /^[*\u2020\u2021\u00a7#\u00b9\u00b2\u00b3\u2070-\u2079]/;
const NOTE_CONT_RE = /^[^\s:=]{1,15}\s*[=:]\s+\S/;
/** Butir catatan: "1.", "1)", "a.", "(a)", "i.", "*", "•", "-", juga tanpa spasi ("2.dadad"). */
const NOTE_ITEM_RE = /^(?:\(?(?:\d{1,2}|[a-zA-Z]|[ivxIVX]{1,4})[.)](?!\d)|[\u2022\u25CF\u25AA\u2013\-*])\s*\S/;
const NOT_NOTE_RE = /^(tabel|table|gambar|figure)\b/i;

/** head = "Keterangan:", item = butir bertanda/bernomor, plain = catatan biasa. */
export type NoteKind = "head" | "item" | "plain";
export type TableNotes = Map<number, NoteKind>;

export type CaptionReport = {
  titled: number;
  moved: number;
  numbered: number;
  refsUpdated: number;
  refsUnresolved: number;
  notes: number;
  uncaptioned: string[];
  appendix: Set<number>;
  regions: VerifyRegion[];
};

/** Run teks pertama paragraf berformat superskrip (penanda catatan a, b, *). */
function startsSuperscript(xml: string): boolean {
  const run = /<w:r[\s>][\s\S]*?<\/w:r>/g;
  let m: RegExpExecArray | null;
  while ((m = run.exec(xml))) {
    if (!/<w:t[\s>]/.test(m[0])) continue;
    return /<w:vertAlign\s+w:val="superscript"/.test(m[0]);
  }
  return false;
}

/** Semua run bertext aslinya lebih kecil dari teks isi. */
function originallySmall(xml: string, bodyHalf: number): boolean {
  const sizes = [...xml.matchAll(/<w:r[\s>][\s\S]*?<\/w:r>/g)]
    .filter((m) => /<w:t[\s>]/.test(m[0]))
    .map((m) => Number(/<w:sz\s+w:val="(\d+)"/.exec(m[0])?.[1] ?? 0));
  return sizes.length > 0 && sizes.every((s) => s > 0 && s <= bodyHalf - 2);
}

/**
 * Deteksi catatan di bawah tabel dari dokumen ASLI (sebelum konversi daftar
 * dan sebelum ukuran font ditimpa). Harus dipanggil paling awal.
 */
export function detectTableNotes(
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  bodyHalf: number,
): TableNotes {
  const notes: TableNotes = new Map();
  const role = (k: number) => outline.items[k]?.role;
  const isEmpty = (n: ParaNode | TableNode) =>
    n.kind === "paragraph" && !n.text.trim() && !n.hasDrawing && !n.hasSectPr && !n.hasPageBreak;
  const nextContent = (k: number) => {
    for (let j = k + 1; j < nodes.length; j++) if (!isEmpty(nodes[j])) return nodes[j];
    return null;
  };

  nodes.forEach((t, i) => {
    if (t.kind !== "table") return;
    let k = i + 1;
    while (k < nodes.length && isEmpty(nodes[k])) k++;
    // Keterangan tabel yang (masih) berada di bawah tabel dilewati.
    if (k < nodes.length && role(k) === "table_caption") {
      k++;
      while (k < nodes.length && isEmpty(nodes[k])) k++;
    }

    let headMode = false;
    let count = 0;
    for (; k < nodes.length; k++) {
      const n = nodes[k];
      if (n.kind !== "paragraph" || n.hasSectPr || n.hasPageBreak || isEmpty(n)) break;
      const r = role(k);
      if (r !== "body" && r !== "empty") break;

      const txt = n.text.trim();
      if (NOT_NOTE_RE.test(txt)) break;
      const head = NOTE_HEAD_RE.test(txt);
      const symbol = NOTE_SYMBOL_RE.test(txt) || startsSuperscript(n.xml);
      const small = originallySmall(n.xml, bodyHalf);
      const marker = n.hasNumbering || NOTE_ITEM_RE.test(txt);

      // Label untuk tabel BERIKUTNYA (mis. sub-judul kuesioner) atau baris
      // tebal seluruhnya adalah judul, bukan catatan.
      const labelsNext = nextContent(k)?.kind === "table" && !head && !symbol && !marker;
      if (labelsNext || (n.allBold && !head)) break;

      const short = txt.length <= 200;
      const ok = count === 0
        ? head || symbol || (small && txt.length <= 300)
        : head || symbol || small || (short && (marker || NOTE_CONT_RE.test(txt) || headMode));
      if (!ok) break;

      const kind: NoteKind = head && (/:$/.test(txt) || txt.length <= 40) ? "head" : marker || symbol ? "item" : "plain";
      notes.set(n.id, kind);
      if (kind === "head") headMode = true;
      count++;
    }
  });
  return notes;
}

function setSize(p: ParaNode, half: number) {
  p.xml = p.xml
    .replace(/<w:sz\s+w:val="\d+"\s*\/>/g, `<w:sz w:val="${half}"/>`)
    .replace(/<w:szCs\s+w:val="\d+"\s*\/>/g, `<w:szCs w:val="${half}"/>`);
}

/* ------------------------------------------------------------------ */
/* Keterangan tabel tanpa nomor                                         */
/* ------------------------------------------------------------------ */

const UNNUMBERED_RE = /^(?:tabel|table)\s+(?!(?:[A-Za-z]{1,2}[.\-]?)?\d)\S.{2,150}$/i;
const NOT_CAPTION_RE = /\b(berikut|di\s*atas|di\s*bawah|menunjukkan|menyajikan|memperlihatkan|disajikan|ditampilkan|terlihat|merupakan|adalah|tersebut)\b/i;

/**
 * "Tabel Sebaran populasi dan sampel" tanpa nomor baru dianggap keterangan bila
 * menempel pada tabel: tepat di atasnya, atau tepat di bawah tabel yang belum
 * punya keterangan. Jalankan sebelum detectTableNotes/applyFormat.
 */
export function promoteUnnumberedCaptions(nodes: (ParaNode | TableNode)[], outline: Outline): Set<number> {
  const out = new Set<number>();
  const blank = (k: number) => { const n = nodes[k]; return n.kind === "paragraph" && !n.text.trim() && !n.hasDrawing; };
  const near = (i: number, step: 1 | -1) => {
    for (let k = i + step; k >= 0 && k < nodes.length; k += step) if (!blank(k)) return k;
    return -1;
  };
  const isTable = (k: number) => k !== -1 && nodes[k].kind === "table";
  const isCaption = (k: number) => k !== -1 && outline.items[k]?.role === "table_caption";

  nodes.forEach((n, i) => {
    const it = outline.items[i];
    if (n.kind !== "paragraph" || !it || it.role !== "body") return;
    if (it.zone !== "main" && it.zone !== "back") return;
    const t = n.text.replace(/\s+/g, " ").trim().replace(/[.:]$/, "");
    if (t.length > 160 || !UNNUMBERED_RE.test(t) || NOT_CAPTION_RE.test(t)) return;
    if (/[.!?;]\s/.test(t) || t.split(" ").length > 18) return;
    const prev = near(i, -1);
    const above = isTable(near(i, 1));
    const below = isTable(prev) && !isCaption(near(prev, -1));
    if (!above && !below) return;
    (it as { role: string }).role = "table_caption";
    out.add(n.id);
  });
  return out;
}

/* ------------------------------------------------------------------ */
/* Penomoran caption (gaya Insert Caption Word) + rujukan di teks       */
/* ------------------------------------------------------------------ */

const NUM = String.raw`(?:[A-Za-z]{1,2}[.\-]?)?\d+(?:[.\-]\d+)*`;
const LABEL_FULL_RE = new RegExp(String.raw`^[ \t\u00a0]*(?:tabel|table)[ \t\u00a0]*\.?[ \t\u00a0]*(?:${NUM})?[ \t\u00a0]*[.:]?[ \t\u00a0]*`, "i");
const LABEL_TEXT_RE = new RegExp(String.raw`^\s*(?:tabel|table)\s*\.?\s*(?:${NUM})?\s*[.:]?\s*`, "i");
const OLD_NUM_RE = /^\s*(?:tabel|table)\s*\.?\s*(\d+(?:[.\-]\d+)*)/i;
/** Hanya "Tabel" berhuruf kapital: "t tabel 1,985" (statistik) tidak ikut. */
const REF_RE = /\b(Tabel|TABEL|Table)(\s+)(\d+(?:[.\-]\d+)*)(?![.,]?\d)/g;
const REF_STRIP_RE = /\b(?:tabel|table)\s+\d+(?:[.\-]\d+)*/gi;
const STOP = new Set(["dan", "atau", "yang", "pada", "untuk", "dengan", "dari", "tiap", "per", "the", "and", "of"]);

const decodeXml = (s: string) =>
  s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const fullText = (xml: string) =>
  [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => decodeXml(m[1])).join("");
const words = (t: string) => ` ${t.toLowerCase().replace(/[^a-z0-9\u00c0-\u024f]+/g, " ").trim()} `;
const keywords = (desc: string) =>
  words(desc).trim().split(" ").filter((w) => w.length > 2 && !STOP.has(w)).slice(0, 3);
/** Kalimat menyebut isi tabelnya ("Ringkasan model pada Tabel 4.11"). */
function mentions(text: string, desc: string): boolean {
  const kw = keywords(desc);
  if (kw.length < 2) return false;
  const hay = words(text);
  return kw.every((w) => hay.includes(` ${w} `));
}

function fromRoman(s: string): number {
  const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
  const a = s.toLowerCase();
  let t = 0;
  for (let k = 0; k < a.length; k++) { const c = v[a[k]]; const n = v[a[k + 1]] ?? 0; t += c < n ? -c : c; }
  return t;
}

/** Nomor bab untuk tiap node di zona isi (null di luar bab). */
function chapterMap(nodes: (ParaNode | TableNode)[], outline: Outline): Map<number, number | null> {
  const out = new Map<number, number | null>();
  let ch: number | null = null;
  let idx = 0;
  nodes.forEach((n, i) => {
    const it = outline.items[i];
    if (!it) return;
    if (it.zone === "main" && n.kind === "paragraph" &&
      (it.role === "chapter_label" || (it.role === "chapter" && !it.continuesChapterLabel))) {
      idx++;
      const m = /bab\s+([ivxlcdm]+|\d+)/i.exec(n.text);
      ch = m ? (/^\d+$/.test(m[1]) ? Number(m[1]) : fromRoman(m[1])) || idx : idx;
    }
    out.set(n.id, it.zone === "main" ? ch : null);
  });
  return out;
}

/**
 * Lepas field yang instruksinya cocok. keepResult=true: hasil field jadi teks
 * biasa (SEQ lama); false: field dibuang seluruhnya (TC).
 */
export function stripFields(xml: string, instr: RegExp, keepResult: boolean): string {
  const hit = (s: string) => instr.test(decodeXml(s));
  xml = xml.replace(/<w:fldSimple\b[^>]*?\/>/g, (all) => (hit(/w:instr="([^"]*)"/.exec(all)?.[1] ?? "") ? "" : all));
  xml = xml.replace(/<w:fldSimple\b([^>]*[^/])>([\s\S]*?)<\/w:fldSimple>/g, (all, attrs: string, inner: string) =>
    hit(/w:instr="([^"]*)"/.exec(attrs)?.[1] ?? "") ? (keepResult ? inner : "") : all);

  const runs = [...xml.matchAll(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g)];
  type F = { runs: number[]; instr: string; stage: "instr" | "result" };
  const stack: F[] = [];
  const drop = new Set<number>();
  runs.forEach((m, k) => {
    const r = m[0];
    const ch = /<w:fldChar\b[^>]*w:fldCharType="(begin|separate|end)"/.exec(r)?.[1];
    if (ch === "begin") { stack.push({ runs: [k], instr: "", stage: "instr" }); return; }
    const top = stack[stack.length - 1];
    if (!top) return;
    if (ch === "separate") { top.runs.push(k); top.stage = "result"; return; }
    if (ch === "end") {
      top.runs.push(k);
      stack.pop();
      if (hit(top.instr)) top.runs.forEach((x) => drop.add(x));
      return;
    }
    if (top.stage === "instr") {
      top.instr += /<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/.exec(r)?.[1] ?? "";
      top.runs.push(k);
    } else if (!keepResult) top.runs.push(k);
  });
  if (!drop.size) return xml;
  let out = "";
  let last = 0;
  runs.forEach((m, k) => {
    if (!drop.has(k)) return;
    out += xml.slice(last, m.index);
    last = m.index! + m[0].length;
  });
  return out + xml.slice(last);
}

/** Bangun ulang caption: "Tabel <bab>." + {SEQ Tabel} + " " + judul, style Caption. */
function rebuildCaption(p: ParaNode, chapter: number, k: number): boolean {
  const before = paragraphLine(p.xml);
  let xml = stripFields(stripFields(p.xml, /^\s*SEQ\b/i, true), /^\s*TC\b/i, false);
  const lm = LABEL_FULL_RE.exec(leadingRaw(xml, 40));
  if (!lm) return false;
  xml = removeLeading(xml, lm[0].length);

  const rPr = /<w:r(?:\s[^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(xml)?.[1] ?? "";
  const run = (inner: string) => `<w:r>${rPr}${inner}</w:r>`;
  const prefix =
    run(`<w:t xml:space="preserve">Tabel ${chapter}.</w:t>`) +
    run(`<w:fldChar w:fldCharType="begin"/>`) +
    run(`<w:instrText xml:space="preserve"> SEQ Tabel \\* ARABIC${k === 1 ? " \\r 1" : ""} </w:instrText>`) +
    run(`<w:fldChar w:fldCharType="separate"/>`) +
    run(`<w:t>${k}</w:t>`) +
    run(`<w:fldChar w:fldCharType="end"/>`) +
    run(`<w:t xml:space="preserve"> </w:t>`);

  const end = xml.indexOf("</w:pPr>");
  const at = end !== -1 ? end + 8 : xml.indexOf(">") + 1;
  xml = xml.slice(0, at) + prefix + xml.slice(at);
  p.xml = withProps(xml, "w:pPr", { "w:pStyle": `<w:pStyle w:val="Caption"/>` }, PPR_ORDER);
  p.text = `Tabel ${chapter}.${k} ${p.text.replace(LABEL_TEXT_RE, "")}`;
  p.hasField = true;
  return paragraphLine(p.xml) !== before;
}

/** Ganti potongan teks (posisi pada gabungan w:t) tanpa merusak run/format. */
function editText(xml: string, edits: { start: number; end: number; text: string }[]): string {
  const re = /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  const toks = [...xml.matchAll(re)].map((m) => decodeXml(m[2]));
  let off = 0;
  const spans = toks.map((t) => { const s = off; off += t.length; return { s, e: off }; });
  const chars = toks.map((t) => t.split(""));
  for (const ed of edits) {
    spans.forEach((sp, ti) => {
      for (let c = Math.max(ed.start, sp.s); c < Math.min(ed.end, sp.e); c++) chars[ti][c - sp.s] = "";
    });
    const ti = spans.findIndex((sp) => ed.start >= sp.s && ed.start < sp.e);
    if (ti !== -1) chars[ti][ed.start - spans[ti].s] = ed.text;
  }
  let k = 0;
  return xml.replace(re, () => `<w:t xml:space="preserve">${escapeXml(chars[k++].join(""))}</w:t>`);
}

export function arrangeTableCaptions(
  model: DocModel,
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  rules: FormatRules,
  tableNotes: TableNotes,
): CaptionReport {
  const rep: CaptionReport = {
    titled: 0, moved: 0, numbered: 0, refsUpdated: 0, refsUnresolved: 0, notes: 0,
    uncaptioned: [], appendix: new Set(), regions: [],
  };
  const itemOf = new Map<number, OutlineItem>();
  nodes.forEach((n, i) => { if (outline.items[i]) itemOf.set(n.id, outline.items[i]); });
  const roleOf = (n: { kind: string }) => (n.kind === "paragraph" ? itemOf.get((n as ParaNode).id)?.role : undefined);
  const isEmpty = (n: (typeof model.body)[number]) =>
    n.kind === "paragraph" && !n.text.trim() && !n.hasDrawing && !n.hasSectPr && !n.hasPageBreak;
  const nextContent = (from: number, step: 1 | -1) => {
    for (let k = from + step; k >= 0 && k < model.body.length; k += step) if (!isEmpty(model.body[k])) return k;
    return -1;
  };

  // 1. Title Case (baris asli dicatat untuk verifikasi).
  const orig = new Map<ParaNode, { line: string; text: string }>();
  nodes.forEach((n, i) => {
    if (n.kind !== "paragraph" || outline.items[i]?.role !== "table_caption" || outline.items[i]?.zone === "front") return;
    orig.set(n, { line: paragraphLine(n.xml), text: n.text });
    const xml = retext(n.xml, titleCaseCaption);
    if (xml !== n.xml) { n.xml = xml; n.text = titleCaseCaption(n.text); rep.titled++; }
  });

  // 2. Keterangan di bawah tabel → pindah ke atas tabel.
  const moved = new Set<ParaNode>();
  for (let ti = 0; ti < model.body.length; ti++) {
    if (model.body[ti].kind !== "table") continue;
    const prev = nextContent(ti, -1);
    if (prev !== -1 && roleOf(model.body[prev]) === "table_caption") continue;
    const next = nextContent(ti, 1);
    if (next === -1 || roleOf(model.body[next]) !== "table_caption") continue;
    const after = nextContent(next, 1);
    if (after !== -1 && model.body[after].kind === "table") continue;
    const cap = model.body[next] as ParaNode;
    model.body.splice(next, 1);
    model.body.splice(ti, 0, cap);
    moved.add(cap);
    rep.moved++;
    ti++;
  }

  // 3. Tabel di bab yang tetap tanpa keterangan → dilaporkan (judul tidak dikarang).
  const chapterOf = chapterMap(nodes, outline);
  const tableNo = new Map<number, number>();
  model.body.forEach((n, bi) => {
    if (n.kind !== "table" || itemOf.get(n.id)?.zone !== "main") return;
    const c = chapterOf.get(n.id) ?? 0;
    const no = (tableNo.get(c) ?? 0) + 1;
    tableNo.set(c, no);
    const prev = nextContent(bi, -1);
    if (prev !== -1 && (roleOf(model.body[prev]) === "table_caption" || model.body[prev].kind === "table")) return;
    rep.uncaptioned.push(c ? `Bab ${c}, tabel ke-${no}` : `tabel ke-${no} sebelum BAB I`);
  });

  // 4. Penomoran ulang per bab + caption gaya Word (style Caption + SEQ Tabel).
  type Entry = { old: string | null; implied: string | null; next: string; pos: number; chapter: number; desc: string };
  const mapping: Entry[] = [];
  const counters = new Map<number, number>();
  const renumbered = new Set<ParaNode>();
  model.body.forEach((n, bi) => {
    if (n.kind !== "paragraph" || roleOf(n) !== "table_caption") return;
    const c = chapterOf.get(n.id);
    if (c == null) {
      // Lampiran: nomor asli (L3.1, A-1, …) dipertahankan, hanya style diseragamkan.
      if (itemOf.get(n.id)?.zone === "back") {
        n.xml = withProps(n.xml, "w:pPr", { "w:pStyle": `<w:pStyle w:val="Caption"/>` }, PPR_ORDER);
        rep.appendix.add(n.id);
      }
      return;
    }
    const om = OLD_NUM_RE.exec(n.text);
    if (!om) {
      const nx = nextContent(bi, 1);
      const target = nx === -1 ? null : model.body[nx];
      if (!target || !(target.kind === "table" || (target.kind === "paragraph" && target.hasDrawing))) return;
    }
    const k = (counters.get(c) ?? 0) + 1;
    counters.set(c, k);
    mapping.push({
      old: om ? om[1].replace(/-/g, ".") : null, implied: null,
      next: `${c}.${k}`, pos: bi, chapter: c, desc: n.text.replace(LABEL_TEXT_RE, ""),
    });
    if (rebuildCaption(n, c, k)) renumbered.add(n);
    rep.numbered++;
  });

  // Keterangan yang dulu tanpa nomor mengisi celah penomoran lama
  // (4.10 → [tanpa nomor] → 4.12 berarti dulu dirujuk sebagai 4.11).
  const usedOld = new Set(mapping.map((m) => m.old).filter(Boolean) as string[]);
  const bump = (o: string, d: number) => {
    const parts = o.split(".");
    const last = Number(parts.pop()) + d;
    return Number.isFinite(last) && last > 0 ? [...parts, String(last)].join(".") : null;
  };
  mapping.forEach((m, i) => {
    if (m.old) return;
    const prev = mapping.slice(0, i).reverse().find((x) => x.chapter === m.chapter && x.old);
    const next = mapping.slice(i + 1).find((x) => x.chapter === m.chapter && x.old);
    const guess = prev ? bump(prev.old!, 1) : next ? bump(next.old!, -1) : null;
    if (guess && !usedOld.has(guess)) m.implied = guess;
  });

  const resolve = (num: string, bi: number, text: string, ch: number | null): Entry | null => {
    const exact = mapping.filter((x) => x.old === num);
    if (exact.length === 1) return exact[0];
    if (exact.length > 1) {
      // Nomor ganda: utamakan tabel yang isinya disebut di kalimat, lalu yang terdekat sesudahnya.
      const named = exact.filter((x) => mentions(text, x.desc));
      const pool = named.length ? named : exact;
      return pool.find((x) => x.pos > bi) ?? [...pool].reverse().find((x) => x.pos < bi) ?? pool[0];
    }
    const named = mapping.filter((x) => (ch == null || x.chapter === ch) && mentions(text, x.desc));
    if (named.length === 1) return named[0];
    return mapping.find((x) => x.implied === num) ?? null;
  };

  // 5. Rujukan di teks ("pada Tabel 4.12") disesuaikan dengan nomor baru.
  model.body.forEach((n, bi) => {
    if (n.kind !== "paragraph" || roleOf(n) === "table_caption") return;
    const zone = itemOf.get(n.id)?.zone;
    if (zone !== "main" && zone !== "back") return;
    const full = fullText(n.xml);
    if (!/\b(?:Tabel|TABEL|Table)\s+\d/.test(full)) return;
    const edits: { start: number; end: number; text: string }[] = [];
    const repl: (string | null)[] = [];
    for (const m of full.matchAll(REF_RE)) {
      const num = m[3].replace(/-/g, ".");
      const hit = resolve(num, bi, full, chapterOf.get(n.id) ?? null);
      if (!hit) { repl.push(null); if (num.includes(".")) rep.refsUnresolved++; continue; }
      if (hit.next === m[3]) { repl.push(null); continue; }
      repl.push(hit.next);
      const start = m.index! + m[1].length + m[2].length;
      edits.push({ start, end: start + m[3].length, text: hit.next });
    }
    if (!edits.length) return;
    const beforeLine = paragraphLine(n.xml);
    n.xml = editText(n.xml, edits);
    let r = 0;
    n.text = n.text.replace(REF_RE, (all, w: string, sp: string) => {
      const v = repl[r++];
      return v ? `${w}${sp}${v}` : all;
    });
    rep.regions.push({
      beforeLines: [beforeLine].filter(Boolean),
      afterLines: [paragraphLine(n.xml)].filter(Boolean),
      beforeText: full.replace(REF_STRIP_RE, " "),
      afterText: fullText(n.xml).replace(REF_STRIP_RE, " "),
    });
    rep.refsUpdated += edits.length;
  });

  // Region verifikasi untuk caption yang dipindah / dinomori ulang.
  for (const [p, o] of orig) {
    if (!moved.has(p) && !renumbered.has(p)) continue;
    rep.regions.push({
      beforeLines: [o.line].filter(Boolean),
      afterLines: [paragraphLine(p.xml)].filter(Boolean),
      beforeText: o.text.replace(LABEL_TEXT_RE, ""),
      afterText: p.text.replace(LABEL_TEXT_RE, ""),
    });
  }

  // 6. Catatan di bawah tabel → 10 pt, rata kiri; butir diberi hanging indent.
  const noteHalf = rules.table().noteSize;
  const HANG = 284;
  const ordered = model.body.filter((n): n is ParaNode => n.kind === "paragraph" && tableNotes.has(n.id));
  ordered.forEach((p, idx) => {
    const kind = tableNotes.get(p.id)!;
    const pos = model.body.indexOf(p);
    const prevIsNote = idx > 0 && model.body.indexOf(ordered[idx - 1]) === pos - 1;
    const nextIsNote = idx < ordered.length - 1 && model.body.indexOf(ordered[idx + 1]) === pos + 1;
    setSize(p, noteHalf);
    p.xml = withProps(p.xml, "w:pPr", {
      "w:spacing": `<w:spacing w:before="${prevIsNote ? 0 : 60}" w:after="${nextIsNote ? 0 : 240}" w:line="240" w:lineRule="auto"/>`,
      "w:ind": kind === "item"
        ? `<w:ind w:left="${HANG}" w:right="0" w:hanging="${HANG}"/>`
        : `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`,
      "w:jc": `<w:jc w:val="left"/>`,
    }, PPR_ORDER);
    rep.notes++;
  });

  return rep;
}