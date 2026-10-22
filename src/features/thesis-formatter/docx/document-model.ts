/**
 * Model dokumen DOCX.
 *
 * Prinsipnya: pohon ini adalah cerminan 1:1 dari `word/document.xml`. Setiap
 * elemen yang tidak kita pahami tetap disimpan apa adanya sebagai `raw`, dan
 * serialisasi mengembalikannya pada posisi semula. Dengan begitu urutan
 * dokumen, gambar, kotak teks, bookmark, dan field tidak pernah hilang hanya
 * karena formatter tidak mengenalinya.
 */

import { elementEnd, innerXml, openTag, splitElements } from "./xml";

export type ParaNode = {
  kind: "paragraph";
  id: number;
  xml: string;
  text: string;
  /** Paragraf memuat gambar (drawing/pict/object) — tidak boleh diperlakukan sebagai teks biasa. */
  hasDrawing: boolean;
  /** Paragraf memuat page break eksplisit di dalam run. */
  hasPageBreak: boolean;
  /** Paragraf sudah punya `w:pageBreakBefore` sendiri sebelum diformat. */
  hadPageBreakBefore: boolean;
  /** Paragraf membawa `w:sectPr` (penanda section break) — tidak boleh dihapus. */
  hasSectPr: boolean;
  /** Paragraf bagian dari daftar bernomor/bullet — indentasinya jangan diganggu. */
  hasNumbering: boolean;
  /** Paragraf memuat field (mis. TOC, nomor halaman). */
  hasField: boolean;
  styleId: string | null;
  /** Perataan yang ditulis langsung di paragraf asli (`w:jc`), bila ada. */
  align: "left" | "center" | "right" | "both" | null;
  /**
   * Jumlah tab di dalam teks. Tab menandakan tata letak manual — kolom tanda
   * tangan, "Nama :\tRaka", bullet ketikan, atau entri daftar isi dengan
   * dot leader. Paragraf seperti ini tidak boleh dijustify atau diindentasi ulang.
   */
  tabCount: number;
  /** Seluruh run bertext dicetak tebal (paragraf isi yang "di-bold semua"). */
  allBold: boolean;
};

export type CellNode = { kind: "cell"; open: string; props: string | null; children: BlockNode[] };
export type RowNode = { kind: "row"; open: string; props: string | null; cells: CellNode[]; extra: string[] };

export type TableNode = {
  kind: "table";
  id: number;
  open: string;
  props: string | null;
  grid: string | null;
  rows: RowNode[];
  /** Elemen lain di dalam `w:tbl` yang dipertahankan apa adanya. */
  extra: string[];
};

export type ContainerNode = { kind: "container"; open: string; close: string; children: BlockNode[] };
export type RawNode = { kind: "raw"; xml: string };

export type BlockNode = ParaNode | TableNode | ContainerNode | RawNode;

export type DocModel = {
  /** XML sebelum `<w:body>` (deklarasi, `w:document`, dst). */
  head: string;
  body: BlockNode[];
  /** `w:sectPr` terakhir milik body (bukan milik paragraf). */
  bodySectPr: string | null;
  tail: string;
};

/**
 * Teks + tab + line break dalam urutan kemunculannya.
 *
 * Tab WAJIB ikut terbaca. Daftar isi Word memakai `<w:tab/>` dengan dot leader
 * di antara judul dan nomor halaman, bukan karakter titik. Tanpa tab,
 * "BAB I PENDAHULUAN<tab>1" terbaca "BAB I PENDAHULUAN1" — lalu dikira judul
 * bab sungguhan dan diberi page break satu per satu.
 *
 * `<w:tab .../>` yang beratribut (definisi tab stop di `w:tabs`) tidak
 * tertangkap karena pola di bawah hanya cocok dengan tab tanpa atribut.
 */
const TEXT_TOKEN_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>|<w:br(?![^>]*w:type="page")[^>]*\/>|<w:cr\s*\/>/g;

function textOf(xml: string): string {
  let out = "";
  let m: RegExpExecArray | null;
  TEXT_TOKEN_RE.lastIndex = 0;
  while ((m = TEXT_TOKEN_RE.exec(xml))) {
    if (m[1] !== undefined) {
      out += m[1]
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, "&");
    } else if (m[0].startsWith("<w:tab")) out += "\t";
    else out += " ";
  }
  // Rapatkan spasi, tetapi JANGAN tab.
  return out.replace(/[ \u00a0]+/g, " ").replace(/ *\t */g, "\t").trim();
}

/** `w:pPr` milik paragraf itu sendiri (anak pertama), bukan yang di dalam w:tc. */
function directProps(el: string, propTag: string): string | null {
  const inner = innerXml(el);
  const trimmed = inner.trimStart();
  if (!trimmed.startsWith(`<${propTag}`)) return null;
  const s = inner.length - trimmed.length;
  return inner.slice(s, elementEnd(inner, s, propTag));
}

function allRunsBold(xml: string): boolean {
  const runs = [...xml.matchAll(/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g)].filter((m) => /<w:t[\s>]/.test(m[1]));
  if (!runs.length) return false;
  return runs.every((m) => /<w:b(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(m[1].split("</w:rPr>")[0] ?? ""));
}

/**
 * Tandai paragraf yang penomorannya diwarisi dari STYLE (mis. "List Number"),
 * bukan dari `w:numPr` di paragrafnya sendiri. Tanpa ini, indentasi daftar
 * milik style tertimpa indentasi paragraf biasa dan nomornya berantakan.
 */
export function markStyleNumbering(nodes: (ParaNode | TableNode)[], stylesXml: string): number {
  const numbered = new Set<string>();
  const basedOn = new Map<string, string>();
  for (const m of stylesXml.matchAll(/<w:style\b[^>]*w:type="paragraph"[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)) {
    if (/<w:numPr>[\s\S]*?<w:numId\s+w:val="(?!0")\d+"/.test(m[2])) numbered.add(m[1]);
    const b = /<w:basedOn\s+w:val="([^"]+)"/.exec(m[2])?.[1];
    if (b) basedOn.set(m[1], b);
  }
  const isNumbered = (id: string, depth = 0): boolean =>
    numbered.has(id) || (depth < 10 && basedOn.has(id) && isNumbered(basedOn.get(id)!, depth + 1));

  let count = 0;
  for (const n of nodes) {
    if (n.kind !== "paragraph" || n.hasNumbering || !n.styleId) continue;
    // numId 0 di paragraf = penomoran style sengaja dimatikan.
    if (/<w:numId\s+w:val="0"/.test(n.xml)) continue;
    if (isNumbered(n.styleId)) { n.hasNumbering = true; count++; }
  }
  return count;
}

function makeParagraph(xml: string, id: number): ParaNode {
  const pPr = directProps(xml, "w:pPr");
  const text = textOf(xml);
  // `w:jc` milik paragraf sendiri — abaikan yang ada di dalam pPrChange/rPr.
  const pPrOwn = pPr ? pPr.replace(/<w:pPrChange[\s\S]*?<\/w:pPrChange>/g, "") : "";
  const jcVal = /<w:jc\s+w:val="([^"]+)"/.exec(pPrOwn)?.[1] ?? null;
  const align =
    jcVal === "center" ? "center"
      : jcVal === "right" || jcVal === "end" ? "right"
      : jcVal === "both" || jcVal === "distribute" ? "both"
      : jcVal === "left" || jcVal === "start" ? "left"
      : null;
  return {
    kind: "paragraph",
    id,
    xml,
    text,
    align,
    tabCount: (text.match(/\t/g) || []).length,
    allBold: allRunsBold(xml),
    hasDrawing: /<w:drawing[\s>]|<w:pict[\s>]|<w:object[\s>]|<v:shape[\s>]/.test(xml),
    hasPageBreak: /<w:br[^>]*w:type="page"/.test(xml),
    hadPageBreakBefore: pPr ? /<w:pageBreakBefore(?:\s[^>]*)?\/>/.test(pPr) && !/w:val="(0|false)"/.test(pPr) : false,
    hasSectPr: pPr ? /<w:sectPr[\s>]/.test(pPr) : false,
    hasNumbering: pPr ? /<w:numPr[\s>]/.test(pPr) : false,
    hasField: /<w:fldChar[\s>]|<w:instrText[\s>]|<w:fldSimple[\s>]/.test(xml),
    styleId: pPr ? /<w:pStyle\s+w:val="([^"]*)"/.exec(pPr)?.[1] ?? null : null,
  };
}

function parseCell(xml: string, ctr: { n: number }): CellNode {
  const children: BlockNode[] = [];
  let props: string | null = null;
  for (const c of splitElements(innerXml(xml))) {
    if (c.tag === "w:tcPr" && props === null) props = c.xml;
    else children.push(parseBlock(c, ctr));
  }
  return { kind: "cell", open: openTag(xml), props, children };
}

function parseRow(xml: string, ctr: { n: number }): RowNode {
  const cells: CellNode[] = [];
  const extra: string[] = [];
  let props: string | null = null;
  for (const c of splitElements(innerXml(xml))) {
    if (c.tag === "w:trPr" && props === null) props = c.xml;
    else if (c.tag === "w:tc") cells.push(parseCell(c.xml, ctr));
    else extra.push(c.xml);
  }
  return { kind: "row", open: openTag(xml), props, cells, extra };
}

function parseTable(xml: string, ctr: { n: number }): TableNode {
  const rows: RowNode[] = [];
  const extra: string[] = [];
  let props: string | null = null;
  let grid: string | null = null;
  for (const c of splitElements(innerXml(xml))) {
    if (c.tag === "w:tblPr" && props === null) props = c.xml;
    else if (c.tag === "w:tblGrid" && grid === null) grid = c.xml;
    else if (c.tag === "w:tr") rows.push(parseRow(c.xml, ctr));
    else extra.push(c.xml);
  }
  return { kind: "table", id: ctr.n++, open: openTag(xml), props, grid, rows, extra };
}

function parseBlock(el: { tag: string; xml: string }, ctr: { n: number }): BlockNode {
  if (el.tag === "w:p") return makeParagraph(el.xml, ctr.n++);
  if (el.tag === "w:tbl") return parseTable(el.xml, ctr);
  if (el.tag === "w:sdt") {
    // Content control (TOC lama, cover page control, dst). Isinya tetap kita
    // format, tapi pembungkusnya dipertahankan persis.
    const inner = innerXml(el.xml);
    const parts = splitElements(inner);
    const idx = parts.findIndex((p) => p.tag === "w:sdtContent");
    if (idx === -1) return { kind: "raw", xml: el.xml };
    const content = parts[idx].xml;
    const before = parts.slice(0, idx).map((p) => p.xml).join("");
    const after = parts.slice(idx + 1).map((p) => p.xml).join("");
    return {
      kind: "container",
      open: `${openTag(el.xml)}${before}${openTag(content)}`,
      close: `</w:sdtContent>${after}</w:sdt>`,
      children: splitElements(innerXml(content)).map((c) => parseBlock(c, ctr)),
    };
  }
  return { kind: "raw", xml: el.xml };
}

export function parseDocument(documentXml: string): DocModel {
  const bodyStart = documentXml.search(/<w:body(?:\s[^>]*)?>/);
  if (bodyStart === -1) throw new Error("word/document.xml tidak memiliki <w:body>");
  const bodyEl = documentXml.slice(bodyStart, elementEnd(documentXml, bodyStart, "w:body"));
  const head = documentXml.slice(0, bodyStart) + openTag(bodyEl);
  const tail = "</w:body>" + documentXml.slice(bodyStart + bodyEl.length);

  const ctr = { n: 0 };
  const children = splitElements(innerXml(bodyEl));
  let bodySectPr: string | null = null;
  const body: BlockNode[] = [];
  for (const c of children) {
    if (c.tag === "w:sectPr") bodySectPr = c.xml;
    else body.push(parseBlock(c, ctr));
  }
  return { head, body, bodySectPr, tail };
}

/** Parse potongan XML blok (paragraf/tabel) menjadi node, untuk elemen buatan formatter. */
export function parseBlocks(xml: string, startId: number): BlockNode[] {
  const ctr = { n: startId };
  return splitElements(xml).map((c) => parseBlock(c, ctr));
}

/* ------------------------------------------------------------------ */

function serializeCell(c: CellNode): string {
  return `${c.open}${c.props ?? ""}${c.children.map(serializeBlock).join("")}</w:tc>`;
}
function serializeRow(r: RowNode): string {
  return `${r.open}${r.props ?? ""}${r.cells.map(serializeCell).join("")}${r.extra.join("")}</w:tr>`;
}
function serializeTable(t: TableNode): string {
  return `${t.open}${t.props ?? ""}${t.grid ?? ""}${t.rows.map(serializeRow).join("")}${t.extra.join("")}</w:tbl>`;
}

export function serializeBlock(n: BlockNode): string {
  switch (n.kind) {
    case "paragraph": return n.xml;
    case "table": return serializeTable(n);
    case "container": return `${n.open}${n.children.map(serializeBlock).join("")}${n.close}`;
    default: return n.xml;
  }
}

export function serializeDocument(m: DocModel): string {
  return `${m.head}${m.body.map(serializeBlock).join("")}${m.bodySectPr ?? ""}${m.tail}`;
}

/* ------------------------------------------------------------------ */

/** Semua paragraf & tabel dalam urutan dokumen, termasuk yang di dalam sdt. */
export function flatten(nodes: BlockNode[], out: (ParaNode | TableNode)[] = []): (ParaNode | TableNode)[] {
  for (const n of nodes) {
    if (n.kind === "paragraph" || n.kind === "table") out.push(n);
    if (n.kind === "container") flatten(n.children, out);
  }
  return out;
}

/** Paragraf di dalam sel tabel (untuk dirapikan terpisah dari body). */
export function cellParagraphs(c: CellNode, out: ParaNode[] = []): ParaNode[] {
  for (const n of c.children) {
    if (n.kind === "paragraph") out.push(n);
    else if (n.kind === "container") flatten(n.children).forEach((x) => x.kind === "paragraph" && out.push(x));
    else if (n.kind === "table") n.rows.forEach((r) => r.cells.forEach((cc) => cellParagraphs(cc, out)));
  }
  return out;
}

export function tableText(t: TableNode): string[][] {
  return t.rows.map((r) => r.cells.map((c) => cellParagraphs(c).map((p) => p.text).join(" ").trim()));
}