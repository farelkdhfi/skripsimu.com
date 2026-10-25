/**
 * Segala hal seputar tabel:
 *  - tata letak tabel (lebar, border, header, perataan sel)
 *  - keterangan (judul) tabel: rata kiri di atas tabel, Title Case, tidak tebal
 *  - catatan di bawah tabel (Sumber/Keterangan/Catatan): 10 pt
 */
import { escapeXml, mergeProps, withProps, PPR_ORDER, TBLPR_ORDER, TCPR_ORDER, TRPR_ORDER } from "./xml";
import { cellParagraphs, tableText, type CellNode, type DocModel, type ParaNode, type TableNode } from "./document-model";
import { formatParagraph, type ApplyReport } from "./apply-format";
import type { FormatRules, ParaFormat, RunFormat, TableFormat } from "../rules/format-rules";
import type { Outline } from "../structure/outline";
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
const LABEL_RE = /^(tabel|table|tbl\.?|gambar|gbr\.?|ilustrasi|figure|fig\.?|grafik|bagan|diagram)\s*[.:]?\s*(?:[A-Za-z]{1,2}[.\-]?)?\d+(?:\s?[.\-]\s?\d+)*\s*[.:\u2013\u2014-]?\s*/i;

/** "TABEL 3.5 KRITERIA RATA-RATA SKOR" → "Tabel 3.5 Kriteria Rata-Rata Skor". */
/** Title Case umum: kata sambung kecil, akronim dipertahankan kecuali teks kapital semua. */
/** Frasa Latin baku: hurufnya dipertahankan seperti aslinya ("In vitro", "et al."). */

/**
 * Title Case judul. Aturannya berdasarkan gaya penulis, bukan daftar kata:
 *  - kapital semua / huruf kecil semua / gaya kalimat → Title Case penuh;
 *  - sebagian besar kata sudah berhuruf awal kapital → huruf kecil yang tersisa dianggap
 *    disengaja (nama spesies, "in vitro", satuan, "pH") dan TIDAK dikapitalkan; hanya kata
 *    tugas ("Dalam", "Terhadap") yang dikecilkan.
 */
export function titleCaseText(s: string, allCapsHint?: boolean): string {
  const letters = s.replace(/[^A-Za-z]/g, "");
  const allCaps = allCapsHint ?? (letters.length > 3 && letters === letters.toUpperCase());
  const words = s.match(WORD_RE) ?? [];
  const content = words.slice(1).filter((w) => !SMALL_WORDS.has(w.toLowerCase()));
  const lowered = content.filter((w) => w[0] !== w[0].toUpperCase()).length;
  const respectLower = !allCaps && content.length > 0 && lowered / content.length < 0.6;
  let first = true;
  return s.replace(WORD_RE, (w) => {
    const isFirst = first;
    first = false;
    if (!allCaps && w.length > 1 && (w === w.toUpperCase() || /[A-Z]/.test(w.slice(1)))) return w;
    const lower = w.toLowerCase();
    if (!isFirst && SMALL_WORDS.has(lower)) return lower;
    if (respectLower && w === lower) return w;
    return lower[0].toUpperCase() + lower.slice(1);
  });
}

/** "GAMBAR 4.2 DISTRIBUSI RESPONDEN" → "Gambar 4.2 Distribusi Responden". */
export function titleCaseCaption(s: string): string {
  const letters = s.replace(/[^A-Za-z]/g, "");
  const allCaps = letters.length > 3 && letters === letters.toUpperCase();
  const label = LABEL_RE.exec(s);
  const head = label ? label[0].replace(/^[a-z]+/i, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase()) : "";
  return head + titleCaseText(s.slice(label ? label[0].length : 0), allCaps);
}

const decode = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

/** Per `w:t`: teks cetak miring (istilah asing / nama ilmiah) hurufnya tidak diubah. */
function italicMask(xml: string, count: number): boolean[] {
  const mask: boolean[] = [];
  for (const m of xml.matchAll(/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g)) {
    const rPr = /^\s*<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(m[1])?.[1] ?? "";
    const it = /<w:i(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(rPr);
    for (const _ of m[1].matchAll(/<w:t(?:\s[^>]*)?>/g)) mask.push(it);
  }
  if (mask.length !== count || mask.every(Boolean)) return new Array(count).fill(false);
  return mask;
}

/** Terapkan transformasi huruf ke teks paragraf tanpa merusak run/format. */
export function retext(xml: string, fn: (s: string) => string): string {
  const re = /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  const texts = [...xml.matchAll(re)].map((m) => decode(m[2]));
  const full = texts.join("");
  const out = fn(full);
  if (out === full || out.length !== full.length) return xml;
  const keep = italicMask(xml, texts.length);
  let pos = 0;
  let k = 0;
  return xml.replace(re, (_m, attr) => {
    const t = texts[k];
    const piece = keep[k] ? t : out.slice(pos, pos + t.length);
    k++;
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

      // Kalimat isi yang panjang (atau beberapa kalimat) bukan catatan, walau hurufnya kecil.
      if (!head && !symbol && (txt.length > 300 || (txt.length > 160 && /[.!?]\s+\S/.test(txt)))) break;
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
/* Catatan di bawah tabel                                               */
/* ------------------------------------------------------------------ */

/** Catatan di bawah tabel → 10 pt, rata kiri; butir diberi hanging indent. */
export function formatTableNotes(model: DocModel, rules: FormatRules, tableNotes: TableNotes): number {
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
  });
  return ordered.length;
}

/**
 * Paragraf isi tepat setelah tabel diberi jarak 12 pt dari garis bawah tabel.
 * Dilewati bila yang menyusul catatan tabel (sudah berjarak), keterangan tabel
 * berikutnya (sudah berjarak), atau baris kosong.
 */
export function spaceAfterTables(model: DocModel, tableNotes: TableNotes): number {
  let count = 0;
  model.body.forEach((t, i) => {
    if (t.kind !== "table") return;
    const p = model.body[i + 1];
    if (!p || p.kind !== "paragraph") return;
    if (tableNotes.has(p.id) || !p.text.trim()) return;
    if (/<w:pStyle w:val="Caption"\/>/.test(p.xml)) return;
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(p.xml)?.[0] ?? "";
    const sp = /<w:spacing\b[^>]*\/>/.exec(pPr)?.[0] ?? "";
    if (Number(/w:before="(\d+)"/.exec(sp)?.[1] ?? 0) >= 240) return;
    const next = sp
      ? sp.replace(/\s*w:before(?:Autospacing|Lines)?="[^"]*"/g, "").replace(/^<w:spacing/, `<w:spacing w:before="240"`)
      : `<w:spacing w:before="240"/>`;
    p.xml = withProps(p.xml, "w:pPr", { "w:spacing": next }, PPR_ORDER);
    count++;
  });
  return count;
}

/**
 * Tabel yang muat dalam satu halaman tidak boleh terpotong: semua baris kecuali
 * baris terakhir diberi keepNext, baris tidak boleh pecah. Tabel panjang dibiarkan
 * bersambung (baris judul sudah diulang per halaman).
 */
export function keepTablesTogether(
  model: DocModel,
  opts: { usableWidth: number; pageHeight: number; halfPts: number },
): number {
  let count = 0;
  const charTw = opts.halfPts * 5;          // perkiraan lebar rata-rata satu karakter
  const lineTw = opts.halfPts * 12;         // tinggi satu baris spasi tunggal + ruang sel
  for (const t of model.body) {
    if (t.kind !== "table" || t.rows.length < 2) continue;
    const texts = tableText(t);
    const cols = Math.max(1, ...texts.map((r) => r.length));
    const cpl = Math.max(4, Math.floor((opts.usableWidth / cols - 200) / charTw));
    let height = 0;
    for (const r of texts) {
      const lines = Math.max(1, ...r.map((c) => Math.ceil(Math.max(1, c.length) / cpl)));
      height += lines * lineTw + 80;
    }
    if (height > opts.pageHeight * 0.8) continue;

    t.rows.forEach((row, ri) => {
      row.props = `<w:trPr>${mergeProps(row.props, { "w:cantSplit": "<w:cantSplit/>" }, TRPR_ORDER).join("")}</w:trPr>`;
      if (ri === t.rows.length - 1) return;
      for (const cell of row.cells) {
        for (const p of cellParagraphs(cell)) {
          p.xml = withProps(p.xml, "w:pPr", { "w:keepNext": "<w:keepNext/>" }, PPR_ORDER);
        }
      }
    });
    count++;
  }
  return count;
}