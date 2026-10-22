/**
 * Low-level OOXML helpers.
 *
 * Semua operasi di sini bersifat bedah (surgical): kita hanya menyentuh elemen
 * properti (pPr/rPr/tblPr/...) dan tidak pernah membuang konten. Itulah yang
 * membuat gambar, tabel, dan urutan dokumen asli tetap utuh.
 */

export type XmlEl = { tag: string; xml: string };

/** Escape teks untuk disisipkan sebagai nilai atribut / konten XML. */
export function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Akhir dari tag pembuka yang dimulai di `start` (indeks '<'). */
function openTagEnd(xml: string, start: number): { end: number; selfClosing: boolean } {
  let i = start + 1;
  let quote: string | null = null;
  for (; i < xml.length; i++) {
    const ch = xml[i];
    if (quote) {
      if (ch === quote) quote = null;
      continue;
    }
    if (ch === '"' || ch === "'") quote = ch;
    else if (ch === ">") return { end: i + 1, selfClosing: xml[i - 1] === "/" };
  }
  return { end: xml.length, selfClosing: true };
}

/**
 * Indeks tepat setelah `</tag>` penutup untuk elemen yang dimulai di `start`.
 * Sadar-kedalaman, sehingga `w:pPr` di dalam `w:pPrChange` tidak membingungkan.
 */
export function elementEnd(xml: string, start: number, tag: string): number {
  const first = openTagEnd(xml, start);
  if (first.selfClosing) return first.end;
  const re = new RegExp(`<${tag}(?=[\\s/>])|</${tag}>`, "g");
  re.lastIndex = first.end;
  let depth = 1;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    if (m[0][1] === "/") {
      if (--depth === 0) return m.index + m[0].length;
    } else if (!openTagEnd(xml, m.index).selfClosing) depth++;
  }
  return xml.length;
}

/** Semua elemen anak langsung dari sebuah string XML (urutan dipertahankan). */
export function splitElements(inner: string): XmlEl[] {
  const out: XmlEl[] = [];
  let i = 0;
  while (i < inner.length) {
    const lt = inner.indexOf("<", i);
    if (lt === -1) break;
    if (inner.startsWith("<!--", lt)) {
      const e = inner.indexOf("-->", lt);
      i = e === -1 ? inner.length : e + 3;
      continue;
    }
    if (inner.startsWith("<?", lt)) {
      const e = inner.indexOf("?>", lt);
      i = e === -1 ? inner.length : e + 2;
      continue;
    }
    const nameMatch = /^<([A-Za-z_][\w.:-]*)/.exec(inner.slice(lt, lt + 80));
    if (!nameMatch) {
      i = lt + 1;
      continue;
    }
    const tag = nameMatch[1];
    const end = elementEnd(inner, lt, tag);
    out.push({ tag, xml: inner.slice(lt, end) });
    i = end;
  }
  return out;
}

/** Isi di antara tag pembuka dan penutup dari sebuah elemen tunggal. */
export function innerXml(el: string): string {
  const open = openTagEnd(el, 0);
  if (open.selfClosing) return "";
  const close = el.lastIndexOf("<");
  return el.slice(open.end, close === -1 ? el.length : close);
}

/** Tag pembuka sebuah elemen, apa adanya (atribut ikut terbawa). */
export function openTag(el: string): string {
  const o = openTagEnd(el, 0);
  return o.selfClosing ? el.slice(0, o.end - 2) + ">" : el.slice(0, o.end);
}

export function tagNameOf(el: string): string {
  return /^<([A-Za-z_][\w.:-]*)/.exec(el)?.[1] ?? "";
}

export function getAttr(el: string, name: string): string | null {
  const re = new RegExp(`\\s${name.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\$&")}="([^"]*)"`);
  return re.exec(openTag(el))?.[1] ?? null;
}

/**
 * Ambil anak pertama bernama `tag` dari sebuah elemen induk.
 * Hanya melihat anak langsung, jadi aman terhadap elemen bersarang.
 */
export function childEl(parent: string, tag: string): string | null {
  for (const c of splitElements(innerXml(parent))) if (c.tag === tag) return c.xml;
  return null;
}

/* ------------------------------------------------------------------ */
/* Penyusunan properti yang sadar-urutan                               */
/* ------------------------------------------------------------------ */

/**
 * Skema OOXML mewajibkan anak-anak elemen properti muncul dalam urutan tertentu.
 * Word menolak membuka dokumen (atau diam-diam membuang properti) bila urutannya
 * salah, jadi daftar ini bukan hiasan.
 */
export const PPR_ORDER = [
  "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl",
  "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
  "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
  "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
  "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
  "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange",
];

export const RPR_ORDER = [
  "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike",
  "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid",
  "w:vanish", "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz",
  "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign",
  "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath", "w:rPrChange",
];

export const TBLPR_ORDER = [
  "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize",
  "w:tblStyleColBandSize", "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders",
  "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription",
  "w:tblPrChange",
];

export const TRPR_ORDER = [
  "w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore", "w:wAfter", "w:cantSplit",
  "w:trHeight", "w:tblHeader", "w:tblCellSpacing", "w:jc", "w:hidden", "w:ins", "w:del",
  "w:trPrChange",
];

export const TCPR_ORDER = [
  "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders", "w:shd", "w:noWrap",
  "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark", "w:headers", "w:tcPrChange",
];

export type PropPatch = Record<string, string | null | undefined>;

/**
 * Gabungkan properti yang sudah ada dengan patch kita, lalu serialisasikan
 * sesuai urutan skema.
 *
 * `null` menghapus properti, `undefined` membiarkannya apa adanya, string
 * menggantinya. Properti yang tidak kita kelola (numPr, sectPr, framePr, ...)
 * selalu ikut terbawa — di situlah struktur asli dokumen terjaga.
 */
export function mergeProps(existing: string | null, patch: PropPatch, order: string[]): string[] {
  const map = new Map<string, string>();
  if (existing) for (const c of splitElements(innerXml(existing))) if (!map.has(c.tag)) map.set(c.tag, c.xml);
  for (const [tag, val] of Object.entries(patch)) {
    if (val === undefined) continue;
    if (val === null) map.delete(tag);
    else map.set(tag, val);
  }
  const ranked = [...map.entries()].sort((a, b) => {
    const ia = order.indexOf(a[0]), ib = order.indexOf(b[0]);
    return (ia === -1 ? order.length : ia) - (ib === -1 ? order.length : ib);
  });
  return ranked.map(([, xml]) => xml);
}

/**
 * Sisipkan/ganti elemen properti (`w:pPr`, `w:rPr`, ...) sebagai anak pertama
 * dari `el`. Mengembalikan elemen utuh yang sudah diperbarui.
 */
export function withProps(el: string, propTag: string, patch: PropPatch, order: string[]): string {
  const tag = tagNameOf(el);
  const open = openTag(el);
  let inner = innerXml(el);

  let existing: string | null = null;
  const trimmedStart = inner.length - inner.trimStart().length;
  if (inner.trimStart().startsWith(`<${propTag}`)) {
    const s = trimmedStart;
    const e = elementEnd(inner, s, propTag);
    existing = inner.slice(s, e);
    inner = inner.slice(0, s) + inner.slice(e);
  }

  const children = mergeProps(existing, patch, order);
  const props = children.length ? `<${propTag}>${children.join("")}</${propTag}>` : "";
  return `${open}${props}${inner}</${tag}>`;
}

/* ------------------------------------------------------------------ */
/* Pembentuk properti yang sering dipakai                               */
/* ------------------------------------------------------------------ */

export const jc = (v: "left" | "center" | "right" | "both") => `<w:jc w:val="${v}"/>`;
export const onOff = (tag: string, on: boolean) => `<${tag}${on ? "" : ' w:val="0"'}/>`;

export const spacing = (o: { before?: number; after?: number; line?: number; rule?: "auto" | "exact" | "atLeast" }) => {
  const a: string[] = [];
  if (o.before !== undefined) a.push(`w:before="${Math.round(o.before)}"`);
  if (o.after !== undefined) a.push(`w:after="${Math.round(o.after)}"`);
  if (o.line !== undefined) a.push(`w:line="${Math.round(o.line)}" w:lineRule="${o.rule ?? "auto"}"`);
  return `<w:spacing ${a.join(" ")}/>`;
};

export const indent = (o: { left?: number; right?: number; firstLine?: number; hanging?: number }) => {
  const a: string[] = [];
  if (o.left !== undefined) a.push(`w:left="${Math.round(o.left)}"`);
  if (o.right !== undefined) a.push(`w:right="${Math.round(o.right)}"`);
  if (o.hanging !== undefined) a.push(`w:hanging="${Math.round(o.hanging)}"`);
  else if (o.firstLine !== undefined) a.push(`w:firstLine="${Math.round(o.firstLine)}"`);
  return `<w:ind ${a.join(" ")}/>`;
};

export const rFonts = (f: string) => {
  const n = escapeXml(f);
  return `<w:rFonts w:ascii="${n}" w:hAnsi="${n}" w:cs="${n}" w:eastAsia="${n}"/>`;
};

export const sz = (halfPoints: number) => `<w:sz w:val="${Math.round(halfPoints)}"/>`;
export const szCs = (halfPoints: number) => `<w:szCs w:val="${Math.round(halfPoints)}"/>`;

export const border = (kind: string, size: number, color = "000000") =>
  `<w:${kind} w:val="single" w:sz="${size}" w:space="0" w:color="${color}"/>`;

export const tblBorders = (size: number, color = "000000") =>
  `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"]
    .map((k) => border(k, size, color))
    .join("")}</w:tblBorders>`;