/**
 * Menerapkan aturan format ke pohon dokumen.
 *
 * Yang ditulis hanya elemen properti (`w:pPr`, `w:rPr`, `w:tblPr`, `w:tcPr`).
 * Tidak ada satu pun run, gambar, sel, atau paragraf yang dibuat, dipindahkan,
 * atau diubah isinya di sini — kecuali pembersihan paragraf kosong yang
 * dinyatakan eksplisit lewat konfigurasi dan selalu dilaporkan.
 */

import {
  elementEnd, indent, jc, mergeProps, onOff, rFonts, spacing, sz, szCs, withProps,
  PPR_ORDER, RPR_ORDER,
} from "./xml";
import {
  cellParagraphs, flatten, type BlockNode, type DocModel,
  type ParaNode,
} from "./document-model";
import type { FormatRules, ParaFormat, RunFormat } from "../rules/format-rules";
import type { Outline } from "../structure/outline";
import type { Role, Zone } from "../structure/roles";
import { normalizeTable } from "./tables";
import { isHeadingRole } from "../structure/roles";


export type ApplyReport = {
  paragraphsFormatted: number;
  runsFormatted: number;
  tablesNormalized: number;
  blankParagraphsRemoved: number;
  duplicatePageBreaksRemoved: number;
  pageBreaksAdded: number;
  /** Judul yang tadinya terpisah dari isinya oleh pindah halaman manual. */
  orphanHeadingsFixed: number;
};

const ALIGN: Record<string, "left" | "center" | "right" | "both"> = {
  left: "left", center: "center", right: "right", both: "both",
};

/** Style judul bawaan Word (juga versi lokal), plus Title/Subtitle. */
const HEADING_STYLE_RE = /^(heading|judul)\s*\d$|^(title|subtitle)$/i;

/* ------------------------------------------------------------------ */
/* Run                                                                  */
/* ------------------------------------------------------------------ */

function rprPatch(f: RunFormat, size = f.size) {
  return {
    "w:rFonts": rFonts(f.font),
    "w:sz": sz(size),
    "w:szCs": szCs(size),
    ...(f.bold === undefined ? {} : { "w:b": onOff("w:b", f.bold), "w:bCs": onOff("w:bCs", f.bold) }),
    // Huruf kapital diterapkan sebagai tampilan (`w:caps`), bukan dengan
    // mengubah teks — sehingga isi dokumen tetap persis seperti aslinya.
    ...(f.caps === undefined ? {} : { "w:caps": f.caps ? "<w:caps/>" : null, "w:smallCaps": f.caps ? undefined : null }),
    ...(f.underline === false ? { "w:u": null } : {}),
    ...(f.italic === false ? { "w:i": null, "w:iCs": null } : {}),
    ...(f.color === undefined ? {} : {
      "w:color": f.color === null ? null : `<w:color w:val="${f.color}"/>`,
      "w:highlight": null,
      "w:shd": null,
    }),
  };
}

/** Ukuran asli sebuah run (half-points), bila ditulis langsung. */
function runSize(run: string): number | null {
  const head = run.slice(0, 600);
  const m = /<w:rPr>[\s\S]*?<w:sz\s+w:val="(\d+)"/.exec(head);
  return m ? Number(m[1]) : null;
}

/** Terapkan format ke seluruh `w:r` pada sebuah string XML (run bersarang dilewati). */
function mapRuns(xml: string, f: RunFormat, onRun: () => void): string {
  const re = /<w:r(?=[\s>/])/g;
  let out = "";
  let i = 0;
  for (; ;) {
    re.lastIndex = i;
    const m = re.exec(xml);
    if (!m) { out += xml.slice(i); break; }
    out += xml.slice(i, m.index);
    const end = elementEnd(xml, m.index, "w:r");
    const run = xml.slice(m.index, end);
    let size = f.size;
    if (f.hierarchy) {
      const orig = runSize(run);
      if (orig !== null && orig > f.hierarchy.threshold) size = f.hierarchy.large;
    }
    out += withProps(run, "w:rPr", rprPatch(f, size), RPR_ORDER);
    onRun();
    i = end;
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* Paragraf                                                             */
/* ------------------------------------------------------------------ */

function pprPatch(fmt: ParaFormat, run: RunFormat, node: ParaNode, forceBreak: boolean | null) {
  const patch: Record<string, string | null | undefined> = {};

  if (fmt.styleId !== undefined) patch["w:pStyle"] = fmt.styleId ? `<w:pStyle w:val="${fmt.styleId}"/>` : null;
  // Paragraf yang bukan judul tidak boleh membawa level outline (masuk daftar isi).
  if (fmt.styleId === null) patch["w:outlineLvl"] = null;
  if (fmt.align) patch["w:jc"] = jc(ALIGN[fmt.align]);
  if (fmt.keepNext !== undefined) patch["w:keepNext"] = fmt.keepNext ? "<w:keepNext/>" : null;
  if (fmt.keepLines !== undefined) patch["w:keepLines"] = fmt.keepLines ? "<w:keepLines/>" : null;

  if (forceBreak !== null) patch["w:pageBreakBefore"] = forceBreak ? "<w:pageBreakBefore/>" : null;

  // Jarak hanya ditimpa bila aturan menyebutkannya. Atribut yang tidak
  // disebut diambil dari paragraf asli, sehingga mis. jarak vertikal sampul
  // tetap utuh walau spasi barisnya diseragamkan.
  if (fmt.before !== undefined || fmt.after !== undefined || fmt.line !== undefined) {
    const old = existingAttrs(node.xml, "w:spacing");
    patch["w:spacing"] = spacing({
      before: fmt.before ?? numAttr(old, "w:before"),
      after: fmt.after ?? numAttr(old, "w:after"),
      line: fmt.line ?? numAttr(old, "w:line"),
      rule: fmt.line !== undefined
        ? "auto"
        : (/w:lineRule="(auto|exact|atLeast)"/.exec(old)?.[1] as "auto" | "exact" | "atLeast" | undefined) ?? "auto",
    });
  }

  // Paragraf berdaftar (bullet/numbering) memakai indentasi dari definisi
  // daftarnya; menimpanya akan merusak perataan nomor.
  if (!fmt.keepIndent && !node.hasNumbering) {
    if (fmt.hanging !== undefined || fmt.left !== undefined || fmt.firstLine !== undefined || fmt.right !== undefined) {
      const old = existingAttrs(node.xml, "w:ind");
      patch["w:ind"] = indent({
        left: fmt.left ?? numAttr(old, "w:left") ?? numAttr(old, "w:start"),
        right: fmt.right ?? numAttr(old, "w:right") ?? numAttr(old, "w:end"),
        hanging: fmt.hanging,
        firstLine: fmt.hanging === undefined ? fmt.firstLine ?? numAttr(old, "w:firstLine") : undefined,
      });
    }
  }

  if (fmt.tocTab !== undefined) patch["w:tabs"] = tocTabs(node.xml, fmt.tocTab);

  // rPr penanda paragraf ikut disamakan agar tinggi baris tidak melar.
  const markRpr = mergeProps(null, rprPatch(run), RPR_ORDER);
  patch["w:rPr"] = `<w:rPr>${markRpr.join("")}</w:rPr>`;

  return patch;
}

/**
 * Tab stop entri daftar: tab kiri (indentasi hierarki 1.1, 1.1.1) dipertahankan,
 * tab kanan diganti satu tab kanan berleader titik di tepi margin kanan.
 */
function tocTabs(xml: string, pos: number): string {
  const pPrEnd = xml.indexOf("</w:pPr>");
  const head = pPrEnd === -1 ? "" : xml.slice(0, pPrEnd).replace(/<w:pPrChange[\s\S]*$/, "");
  const old = /<w:tabs>([\s\S]*?)<\/w:tabs>/.exec(head)?.[1] ?? "";
  const kept = [...old.matchAll(/<w:tab\b[^>]*\/>/g)]
    .map((m) => m[0])
    .filter((t) => !/w:val="(right|end|clear)"/.test(t) && Number(/w:pos="(-?\d+)"/.exec(t)?.[1] ?? 0) < pos - 200);
  return `<w:tabs>${kept.join("")}<w:tab w:val="right" w:leader="dot" w:pos="${pos}"/></w:tabs>`;
}

/** Atribut sebuah elemen properti di pPr paragraf asli (mis. `w:spacing`). */
function existingAttrs(xml: string, tag: string): string {
  const pPrEnd = xml.indexOf("</w:pPr>");
  if (pPrEnd === -1) return "";
  const head = xml.slice(0, pPrEnd).replace(/<w:pPrChange[\s\S]*$/, "");
  return new RegExp(`<${tag}(\\s[^>]*)/?>`).exec(head)?.[1] ?? "";
}

function numAttr(attrs: string, name: string): number | undefined {
  const m = new RegExp(`${name}="(-?\\d+)"`).exec(attrs);
  return m ? Number(m[1]) : undefined;
}

/**
 * Buang spasi (dan tab, bila diminta) di TEPI paragraf.
 *
 * Penulis sering memusatkan atau mengindentasi baris dengan mengetik spasi
 * atau tab. Setelah paragraf diberi perataan/indentasi yang benar, spasi itu
 * justru menggeser teks. Hanya tepi yang disentuh; spasi di antara kata tetap.
 */
function trimEdges(xml: string, leadingTabs: boolean): string {
  // Tepi kiri: buang token <w:tab/> dan spasi di awal sampai bertemu huruf.
  let out = xml;
  const tokenRe = /<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>/g;
  let m: RegExpExecArray | null;
  while ((m = tokenRe.exec(out))) {
    if (m[0].startsWith("<w:tab")) {
      if (!leadingTabs) break;
      out = out.slice(0, m.index) + out.slice(m.index + m[0].length);
      tokenRe.lastIndex = m.index;
      continue;
    }
    const trimmed = m[2].replace(/^[ \u00a0]+/, "");
    if (trimmed !== m[2]) {
      const rep = `<w:t${m[1] ?? ""}>${trimmed}</w:t>`;
      out = out.slice(0, m.index) + rep + out.slice(m.index + m[0].length);
      tokenRe.lastIndex = m.index + rep.length;
    }
    if (trimmed) break;
  }
  // Tepi kanan: spasi di akhir teks terakhir.
  const all = [...out.matchAll(/<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g)];
  for (let k = all.length - 1; k >= 0; k--) {
    const t = all[k];
    const trimmed = t[2].replace(/[ \u00a0]+$/, "");
    if (trimmed !== t[2]) {
      const idx = t.index!;
      out = out.slice(0, idx) + `<w:t${t[1] ?? ""}>${trimmed}</w:t>` + out.slice(idx + t[0].length);
    }
    if (trimmed) break;
  }
  return out;
}

export function formatParagraph(
  node: ParaNode, fmt: ParaFormat, run: RunFormat, forceBreak: boolean | null, rep: ApplyReport,
): void {
  let xml = node.xml;
  if (!fmt.keepIndent && !node.hasNumbering) xml = trimEdges(xml, Boolean(fmt.stripLeadingTabs));
  xml = mapRuns(xml, run, () => rep.runsFormatted++);
  xml = withProps(xml, "w:pPr", pprPatch(fmt, run, node, forceBreak), PPR_ORDER);
  node.xml = xml;
  rep.paragraphsFormatted++;
}

/* ------------------------------------------------------------------ */
/* Pembersihan halaman kosong                                           */
/* ------------------------------------------------------------------ */

function isDroppableEmpty(n: BlockNode): n is ParaNode {
  return (
    n.kind === "paragraph" && !n.text.trim() && !n.hasDrawing && !n.hasSectPr &&
    !n.hasField && !n.hasPageBreak && !/<w:bookmarkStart/.test(n.xml)
  );
}

function collapseBlanks(nodes: BlockNode[], zoneOf: (n: BlockNode) => Zone, rep: ApplyReport): BlockNode[] {
  const out: BlockNode[] = [];
  let streak = 0;
  for (const n of nodes) {
    if (n.kind === "container") {
      n.children = collapseBlanks(n.children, zoneOf, rep);
      out.push(n);
      streak = 0;
      continue;
    }
    // Halaman sampul memakai paragraf kosong sebagai penata jarak vertikal —
    // jangan diutak-atik.
    if (isDroppableEmpty(n) && zoneOf(n) !== "cover") {
      streak++;
      if (streak > 1) { rep.blankParagraphsRemoved++; continue; }
    } else streak = 0;
    out.push(n);
  }
  return out;
}

/** Beberapa `<w:br w:type="page"/>` beruntun dalam satu paragraf → satu saja. */
function dedupePageBreaks(node: ParaNode, rep: ApplyReport) {
  const matches = node.xml.match(/<w:br[^>]*w:type="page"[^>]*\/>/g);
  if (!matches || matches.length < 2) return;
  let first = true;
  node.xml = node.xml.replace(/<w:br[^>]*w:type="page"[^>]*\/>/g, (m) => {
    if (first) { first = false; return m; }
    rep.duplicatePageBreaksRemoved++;
    return "";
  });
}

/* ------------------------------------------------------------------ */
/* Entry point                                                          */
/* ------------------------------------------------------------------ */

export function applyFormat(
  model: DocModel,
  outline: Outline,
  rules: FormatRules,
  opts: { cleanupBlankPages: boolean; normalizeTables: boolean },
): ApplyReport {
  const rep: ApplyReport = {
    paragraphsFormatted: 0, runsFormatted: 0, tablesNormalized: 0,
    blankParagraphsRemoved: 0, duplicatePageBreaksRemoved: 0, pageBreaksAdded: 0, orphanHeadingsFixed: 0,
  };

  const nodes = flatten(model.body);
  const zoneById = new Map<number, Zone>();
  nodes.forEach((n, i) => zoneById.set(n.id, outline.items[i]?.zone ?? "main"));

  const page = rules.page("main");
  const usableWidth = page.widthTwips - page.margin.left - page.margin.right;
  const tableFmt = rules.table();

  nodes.forEach((node, i) => {
    const item = outline.items[i];
    if (!item) return;
    const { role, zone } = item;

    if (node.kind === "table") {
      if (opts.normalizeTables) {
        const texts = node.rows.map((r) => r.cells.map((c) => cellParagraphs(c).map((p) => p.text).join(" ").trim()));
        normalizeTable(node, { fmt: tableFmt, rules, usableWidth, texts, rep }, rep);
      }
      return;
    }

    if (opts.cleanupBlankPages) dedupePageBreaks(node, rep);

    let fmt = rules.paragraph(role as Role, zone, item.hint);
    // Teks isi yang terlanjur memakai style Heading → dijadikan teks biasa.
    if (!isHeadingRole(role as Role) && fmt.styleId === undefined && node.styleId && HEADING_STYLE_RE.test(node.styleId)) {
      fmt = { ...fmt, styleId: null };
    }
    const run = rules.run(role as Role, zone, item.hint);

    // Hanya tambahkan page break bila halaman baru belum dimulai oleh
    // dokumen aslinya. Inilah yang mencegah halaman kosong ganda.
    let forceBreak: boolean | null = null;
    if (fmt.pageBreakBefore) {
      forceBreak = !item.breakAlreadyPresent;
      if (forceBreak && !node.hadPageBreakBefore) rep.pageBreaksAdded++;
      // Dokumen sudah memulai halaman baru di sini (break eksplisit atau
      // section break). Menyisakan pageBreakBefore akan membuat halaman kosong.
      if (!forceBreak && node.hadPageBreakBefore) rep.duplicatePageBreaksRemoved++;
    }

    formatParagraph(node, fmt, run, forceBreak, rep);
  });

  if (opts.cleanupBlankPages) {
    model.body = collapseBlanks(model.body, (n) => (n.kind === "paragraph" ? zoneById.get(n.id) ?? "main" : "main"), rep);
  }

  return rep;
}

/* ------------------------------------------------------------------ */
/* Halaman kosong akibat pindah halaman beruntun                        */
/* ------------------------------------------------------------------ */

const PBB_RE = /<w:pageBreakBefore(?:\s+w:val="(?:1|true|on)")?\s*\/>/;
const PAGE_BR_RE = /<w:br\s[^>]*w:type="page"[^>]*\/>/g;

type Boundary =
  | { kind: "start" }
  | { kind: "pbb"; node: ParaNode }          // w:pageBreakBefore di pPr
  | { kind: "br"; node: ParaNode }           // <w:br w:type="page"/> di paragraf tanpa teks
  | { kind: "fixed" };                       // section break / br di paragraf bertext

function hasContent(n: ParaNode): boolean {
  return Boolean(n.text.trim()) || n.hasDrawing;
}

function sectionBreaks(n: ParaNode): boolean {
  if (!n.hasSectPr) return false;
  const s = /<w:sectPr[\s\S]*?<\/w:sectPr>/.exec(n.xml)?.[0] ?? "";
  return !/<w:type\s+w:val="continuous"/.test(s);
}

/**
 * Hapus pindah halaman yang tidak memisahkan isi apa pun.
 *
 * Dokumen nyata sering memuat page break manual, lalu section break, lalu
 * judul yang juga diberi "page break before" — tiga batas halaman berturut
 * tanpa isi di antaranya, hasilnya halaman kosong. Fungsi ini menyusuri batas
 * halaman secara berurutan; bila dua batas bertemu tanpa isi di antaranya,
 * yang bisa dihapus dihapus. Section break tidak pernah dihapus karena ia
 * membawa ukuran kertas, orientasi, dan header/footer.
 */
export function collapseEmptyPages(model: DocModel, rep: ApplyReport, headingIds: Set<number> = new Set()): void {
  let last: Boundary = { kind: "start" };
  let content = false;
  /** Isi terakhir sebelum batas ini adalah judul (bab/subbab). */
  let lastWasHeading = false;
  const drop = new Set<ParaNode>();

  const remove = (b: Boundary): boolean => {
    if (b.kind === "pbb") {
      b.node.xml = b.node.xml.replace(PBB_RE, "");
      b.node.hadPageBreakBefore = false;
      rep.duplicatePageBreaksRemoved++;
      return true;
    }
    if (b.kind === "br") {
      b.node.xml = b.node.xml.replace(PAGE_BR_RE, "");
      b.node.hasPageBreak = false;
      if (isDroppableEmpty(b.node)) drop.add(b.node);
      rep.duplicatePageBreaksRemoved++;
      return true;
    }
    return false;
  };

  const boundary = (b: Boundary) => {
    // Pindah halaman manual tepat setelah judul memisahkan judul dari isinya:
    // judul tertinggal sendirian di bawah halaman. Buang pindah halamannya.
    // Pengecualian: pageBreakBefore milik judul berikutnya (mis. BAB baru) sah.
    if (content && lastWasHeading && (b.kind === "br" || (b.kind === "pbb" && !headingIds.has(b.node.id)))) {
      remove(b);
      rep.orphanHeadingsFixed++;
      return;
    }
    if (content) { last = b; content = false; return; }
    // Dua batas halaman tanpa isi di antaranya → satu harus pergi.
    // Utamakan menghapus yang LEBIH AWAL (biasanya page break manual sisa
    // penulis), kecuali yang awal tidak bisa dihapus.
    if (last.kind === "start") { if (!remove(b)) last = b; return; }
    if (remove(last)) { last = b; return; }
    if (!remove(b)) last = b;
  };

  const markContent = (isHeading: boolean) => { content = true; lastWasHeading = isHeading; };

  for (const n of model.body) {
    if (n.kind === "table") { markContent(false); continue; }
    if (n.kind === "container") {
      if (flatten(n.children).some((x) => x.kind === "table" || hasContent(x as ParaNode))) markContent(false);
      continue;
    }
    if (n.kind !== "paragraph") continue;

    const isHeading = headingIds.has(n.id);
    const pPrEnd = n.xml.indexOf("</w:pPr>");
    if (pPrEnd !== -1 && PBB_RE.test(n.xml.slice(0, pPrEnd))) boundary({ kind: "pbb", node: n });

    PAGE_BR_RE.lastIndex = 0;
    const brIdx = n.xml.search(PAGE_BR_RE);
    if (brIdx !== -1) {
      if (!hasContent(n)) boundary({ kind: "br", node: n });
      else {
        // Break di depan teks: batas dulu baru isi. Di belakang: isi dulu.
        const textIdx = n.xml.search(/<w:t[\s>]|<w:drawing[\s>]/);
        if (textIdx !== -1 && brIdx < textIdx) { boundary({ kind: "fixed" }); markContent(isHeading); }
        else { markContent(isHeading); boundary({ kind: "fixed" }); }
      }
    } else if (hasContent(n)) markContent(isHeading);

    if (sectionBreaks(n)) boundary({ kind: "fixed" });
  }

  if (drop.size) model.body = model.body.filter((n) => !(n.kind === "paragraph" && drop.has(n)));
}