/**
 * Blok tanda tangan → tabel tanpa garis dengan ruang tanda tangan yang cukup.
 *
 * Dua sumber blok:
 *  1. Rangkaian paragraf pendek (tanggal, jabatan, materai, garis tanda tangan,
 *     nama, nomor identitas) → disusun jadi tabel tanpa garis.
 *  2. Tabel asli yang isinya tanda tangan ("Pembimbing I | Pembimbing II", nama, NIDN)
 *     → TIDAK diperlakukan sebagai tabel data: garis & arsiran dibuang, isi rata tengah,
 *     baris kosong jadi ruang tanda tangan, baris tidak terpotong halaman.
 *
 * Tata letak paragraf:
 *  - baris bertab (beberapa penanda tangan sejajar, 2–4 kolom) → satu kolom per penanda tangan
 *  - baris rata tengah → satu sel selebar tabel (mis. "Mengetahui," di bawah dua pembimbing)
 *  - baris rata kanan / menjorok kanan / diawali tab → kolom kanan
 *  - ruang tanda tangan ±2,3 cm sebelum nama (atau sebelum garis tanda tangan / mengapit materai)
 * Satu blok (tanggal + tabel + blok "Mengetahui") diikat agar tidak terpisah halaman.
 */
import { childEl, innerXml, openTag, splitElements, withProps, rFonts, sz, szCs, PPR_ORDER, RPR_ORDER } from "./xml";
import { cellParagraphs, parseBlocks, type DocModel, type ParaNode, type TableNode } from "./document-model";
import type { Outline, OutlineItem } from "../structure/outline";
import { paragraphLine } from "./verify";

export type SigRegion = { beforeLines: string[]; afterLines: string[]; beforeText: string; afterText: string };
export type SigReport = {
  blocks: number;
  /** Tabel tanda tangan asli yang dirapikan (garis dibuang). */
  cleaned: number;
  regions: SigRegion[];
  added: { tables: number; rows: number; cells: number };
};

type K = "date" | "role" | "materai" | "sigline" | "id" | "name" | "plain";
type Cls = "multi" | "left" | "center" | "right";
type Line = { p: ParaNode; parts: string[]; kinds: K[]; cls: Cls; gapBefore: boolean };
type Item = { xml: string; kind: K };
type SigFmt = { font: string; halfPts: number };

const DATE_RE = /^[A-Z][A-Za-z.'\- ]{1,40},\s*(\d{1,2}\s+)?[A-Za-z]{3,12}\s+\d{4}\.?$/;
const ROLE_RE = /^(mengetahui|menyetujui|disetujui|disahkan|pembimbing|penguji|ketua|dekan|wakil|rektor|direktur|kepala|sekretaris|koordinator|kaprodi|yang\s+menyatakan|yang\s+membuat|penulis|peneliti|mahasiswa|dosen|pembina|atas\s+nama|a\.n\.|an\.)/i;
const SELF_ROLE_RE = /(yang\s+menyatakan|yang\s+membuat|penulis|peneliti|mahasiswa)/i;
const MATERAI_RE = /^\[?\s*(materai|meterai)/i;
const ID_RE = /^(NIP|NIDN|NIDK|NIK|NIM|NPM|NRP|NISN|NUPTK)\b/i;
const DEGREE_RE = /(\b(Prof|Dr|Drs|Dra|Ir|H|Hj)\.\s)|(,\s*(S|M|Dr|Ph|Sp|Ak|Apt)\.?\s?[A-Za-z]{0,6}\.?)|(\bPh\.?D\b)/;
const PLAIN_RE = /^[A-Z][A-Za-z.'\u2019\-]*( [A-Z][A-Za-z.'\u2019\-]*){1,5}$/;
/** "(Raka Aditya Pramudya)" — nama penanda tangan dalam kurung. */
const PAREN_NAME_RE = /^\(\s*[A-Z][^()]{2,70}\)$/;
/** "R.A.P." — inisial penulis. */
const INITIALS_RE = /^(?:[A-Z]\.){2,5}$/;
/** "(.....................)", "______________" — garis tempat tanda tangan. */
const SIGLINE_RE = /^[([]?[\s._\-\u2013\u2014\u2026]{4,}[)\]]?$/;

const SIG_SPACE = 1300;        // ±2,3 cm ruang tanda tangan
const MATERAI_PAD = 600;       // jarak di atas & bawah tulisan materai
const ROW_GAP = 240;           // setelah baris tanpa penanda tangan (mis. tanggal)
const SIGNER_ROW_GAP = 567;    // ±1 cm setelah baris penanda tangan, sebelum blok berikutnya
const BLOCK_GAP = 567;         // ±1 cm antara teks isi dan blok tanda tangan

const EMPTY_P = `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>`;
const RUN_RE = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g;
const norm = (s: string) => s.replace(/\s+/g, " ").trim();

function classifyPart(t: string): K | null {
  if (!t || t.length > 90) return null;
  if (SIGLINE_RE.test(t)) return "sigline";
  if (ID_RE.test(t)) return "id";
  if (MATERAI_RE.test(t)) return "materai";
  if (DATE_RE.test(t)) return "date";
  if (ROLE_RE.test(t) && t.length <= 80) return "role";
  if (PAREN_NAME_RE.test(t)) return "name";
  if (t.includes(":")) return null;
  if (DEGREE_RE.test(t) && t.length <= 80) return "name";
  if (INITIALS_RE.test(t)) return "plain";
  if (PLAIN_RE.test(t) && t.length <= 60) return "plain";
  return null;
}

function alignClass(p: ParaNode, parts: number): Cls {
  if (parts > 1) return "multi";
  // Baris yang didorong dengan tab ("\t\tMengetahui,") berada di kolom kanan.
  if (/^\t/.test(p.text) && p.text.trim()) return "right";
  const end = p.xml.indexOf("</w:pPr>");
  const pPr = end === -1 ? "" : p.xml.slice(0, end);
  const jc = /<w:jc\s+w:val="(\w+)"/.exec(pPr)?.[1];
  const left = Number(/<w:ind\b[^>]*w:(?:left|start)="(-?\d+)"/.exec(pPr)?.[1] ?? 0);
  if (jc === "right" || jc === "end" || left >= 2000) return "right";
  if (jc === "center") return "center";
  return "left";
}

/** Pecah satu paragraf bertab menjadi beberapa paragraf (satu per kolom). */
function splitColumns(xml: string): string[] {
  const open = openTag(xml);
  const pPrEnd = xml.indexOf("</w:pPr>");
  const pPr = pPrEnd === -1 ? "" : xml.slice(open.length, pPrEnd + 8);
  const inner = innerXml(xml).slice(pPr.length);
  const cols: string[][] = [[]];
  let pendingSep = false;
  const hasText = (c: string[]) => c.some((x) => /<w:t[\s>]/.test(x));

  for (const el of splitElements(inner)) {
    if (el.tag !== "w:r") { cols[cols.length - 1].push(el.xml); continue; }
    const rPr = childEl(el.xml, "w:rPr") ?? "";
    const kids = splitElements(innerXml(el.xml)).filter((k) => k.tag !== "w:rPr");
    let buf: string[] = [];
    const emit = () => { if (buf.length) { cols[cols.length - 1].push(`<w:r>${rPr}${buf.join("")}</w:r>`); buf = []; } };
    for (const k of kids) {
      if (k.tag === "w:tab" || k.tag === "w:ptab") { emit(); pendingSep = true; continue; }
      if (pendingSep) {
        emit();
        if (hasText(cols[cols.length - 1])) cols.push([]);
        pendingSep = false;
      }
      buf.push(k.xml);
    }
    emit();
  }
  return cols.filter(hasText).map((c) => `<w:p>${pPr}${c.join("")}</w:p>`);
}

function stripLeadingTabs(xml: string): string {
  const t = xml.search(/<w:t[\s>]/);
  return t === -1 ? xml : xml.slice(0, t).replace(/<w:tab\s*\/>/g, "") + xml.slice(t);
}

function fmtPara(xml: string, before: number, after: number, keepNext: boolean): string {
  return withProps(stripLeadingTabs(xml), "w:pPr", {
    "w:pStyle": null,
    "w:keepNext": keepNext ? "<w:keepNext/>" : null,
    "w:keepLines": "<w:keepLines/>",
    "w:pageBreakBefore": null,
    "w:numPr": null,
    "w:tabs": null,
    "w:spacing": `<w:spacing w:before="${before}" w:after="${after}" w:line="240" w:lineRule="auto"/>`,
    "w:ind": `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`,
    "w:jc": `<w:jc w:val="center"/>`,
  }, PPR_ORDER);
}

/** Font & ukuran isi tanda tangan diseragamkan (tebal/miring/garis bawah asli dipertahankan). */
function fmtRuns(xml: string, f: SigFmt): string {
  return xml.replace(RUN_RE, (run) =>
    withProps(run, "w:rPr", { "w:rFonts": rFonts(f.font), "w:sz": sz(f.halfPts), "w:szCs": szCs(f.halfPts) }, RPR_ORDER));
}

/** Indeks baris tempat ruang tanda tangan disisipkan (sebelum nama / garis / mengapit materai). */
function signatureSpacing(kinds: K[]): Map<number, { before?: number; after?: number }> {
  const out = new Map<number, { before?: number; after?: number }>();
  let nameIdx = kinds.findIndex((k) => k === "id") - 1;
  if (nameIdx < 0) nameIdx = kinds.findIndex((k) => k === "name");
  if (nameIdx < 0) {
    const m = kinds.findIndex((k) => k === "materai");
    if (m !== -1 && m + 1 < kinds.length) nameIdx = m + 1;
  }
  if (nameIdx < 0) nameIdx = kinds.findIndex((k) => k === "sigline");
  if (nameIdx <= 0) return out;
  const prev = kinds[nameIdx - 1];
  if (prev === "materai") out.set(nameIdx - 1, { before: MATERAI_PAD, after: MATERAI_PAD });
  else if (prev === "sigline") out.set(nameIdx - 1, { before: SIG_SPACE });
  else out.set(nameIdx, { before: SIG_SPACE });
  return out;
}

const isEmptyPara = (p: ParaNode) =>
  !p.text.trim() && !p.hasDrawing && !p.hasSectPr && !p.hasPageBreak && !p.hasField;

/* ------------------------------------------------------------------ */
/* Tabel tanda tangan asli                                               */
/* ------------------------------------------------------------------ */

function isSignatureTable(t: TableNode): boolean {
  if (!t.rows.length || t.rows.length > 10 || t.rows.some((r) => r.cells.length > 4)) return false;
  const lines: string[] = [];
  for (const r of t.rows) {
    for (const c of r.cells) {
      if (c.children.some((ch) => ch.kind === "table")) return false;
      for (const p of cellParagraphs(c)) {
        if (p.hasDrawing) return false;
        const tx = norm(p.text);
        if (tx) lines.push(tx);
      }
    }
  }
  if (lines.length < 2) return false;
  const kinds = lines.map(classifyPart);
  if (kinds.some((k) => k === null)) return false;
  const anchored = kinds.some((k) => k === "name" || k === "id" || k === "sigline");
  const context = kinds.some((k) => k === "role" || k === "date" || k === "materai");
  return anchored && context;
}

/**
 * Tabel yang isinya blok tanda tangan (dipanggil SEBELUM applyFormat, supaya tabel ini
 * tidak dinormalkan sebagai tabel data: garis, baris judul tebal, header berulang).
 */
export function findSignatureTables(nodes: (ParaNode | TableNode)[], outline: Outline): Set<number> {
  const out = new Set<number>();
  nodes.forEach((n, i) => {
    if (n.kind !== "table") return;
    const it = outline.items[i];
    if (!it || it.zone === "cover") return;
    if (isSignatureTable(n)) out.add(n.id);
  });
  return out;
}

function cleanSignatureTable(t: TableNode, f: SigFmt, keepAfter: boolean, width: number): void {
  const nil = (k: string) => `<w:${k} w:val="nil"/>`;
  // Kolom penanda tangan sama lebar, memenuhi lebar teks (simetris).
  const spanOf = (props: string | null) => Number(/<w:gridSpan w:val="(\d+)"/.exec(props ?? "")?.[1] ?? 1);
  const cols = Math.max(1, ...t.rows.map((r) => r.cells.reduce((s, c) => s + spanOf(c.props), 0)));
  const colW = Math.floor(width / cols);
  t.grid = `<w:tblGrid>${`<w:gridCol w:w="${colW}"/>`.repeat(cols)}</w:tblGrid>`;
  t.props =
    `<w:tblPr><w:tblW w:w="${colW * cols}" w:type="dxa"/><w:jc w:val="center"/>` +
    `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map(nil).join("")}</w:tblBorders>` +
    `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="57" w:type="dxa"/><w:right w:w="57" w:type="dxa"/></w:tblCellMar>` +
    `<w:tblLook w:val="0000"/></w:tblPr>`;

  const rowTexts = t.rows.map((r) => r.cells.flatMap((c) => cellParagraphs(c).map((p) => norm(p.text))).filter(Boolean));
  const hasSpaceRow = rowTexts.some((tx) => tx.length === 0);

  t.rows.forEach((row, ri) => {
    const empty = rowTexts[ri].length === 0;
    const lastRow = ri === t.rows.length - 1;
    row.props = `<w:trPr><w:cantSplit/>${empty ? `<w:trHeight w:val="${SIG_SPACE}" w:hRule="atLeast"/>` : ""}</w:trPr>`;
    for (const cell of row.cells) {
      const keep = (re: RegExp) => re.exec(cell.props ?? "")?.[0] ?? "";
      const span = spanOf(cell.props);
      cell.props =
        `<w:tcPr><w:tcW w:w="${colW * span}" w:type="dxa"/>${keep(/<w:gridSpan\b[^>]*\/>/)}${keep(/<w:vMerge\b[^>]*\/>/)}` +
        `<w:tcBorders>${["top", "left", "bottom", "right"].map(nil).join("")}</w:tcBorders><w:vAlign w:val="top"/></w:tcPr>`;

      const ps = cellParagraphs(cell);
      const textual = ps.filter((p) => norm(p.text));
      // Tanpa baris kosong khusus: ruang tanda tangan dibuat di dalam sel (sebelum nama).
      const sp = hasSpaceRow ? new Map<number, { before?: number; after?: number }>()
        : signatureSpacing(textual.map((p) => classifyPart(norm(p.text)) ?? "plain"));
      for (const p of ps) {
        const k = textual.indexOf(p);
        const s = k === -1 ? {} : sp.get(k) ?? {};
        p.xml = fmtRuns(fmtPara(p.xml, s.before ?? 0, s.after ?? 0, !lastRow || keepAfter), f);
      }
    }
  });
}

/* ------------------------------------------------------------------ */
/* Proses utama                                                          */
/* ------------------------------------------------------------------ */

export function buildSignatureTables(
  model: DocModel,
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  opts: { width: number; font: string; halfPts: number; tables?: Set<number> },
): SigReport {
  const rep: SigReport = { blocks: 0, cleaned: 0, regions: [], added: { tables: 0, rows: 0, cells: 0 } };
  const itemOf = new Map<number, OutlineItem>();
  nodes.forEach((n, i) => { if (outline.items[i]) itemOf.set(n.id, outline.items[i]); });
  const fmt: SigFmt = { font: opts.font, halfPts: opts.halfPts };

  const prevContent = (at: number) => {
    for (let k = at - 1; k >= 0; k--) {
      const n = model.body[k];
      if (n.kind === "paragraph" && isEmptyPara(n)) continue;
      return n;
    }
    return null;
  };
  const nextContent = (at: number) => {
    for (let k = at + 1; k < model.body.length; k++) {
      const n = model.body[k];
      if (n.kind === "paragraph" && isEmptyPara(n)) continue;
      return n;
    }
    return null;
  };
  /** Baris tanggal / jabatan tepat di atas blok ikut diikat ke blok. */
  const bindPrevious = (at: number) => {
    const prev = prevContent(at);
    if (prev?.kind !== "paragraph") return;
    const t = norm(prev.text);
    if (DATE_RE.test(t) || (ROLE_RE.test(t) && t.length <= 80)) {
      prev.xml = withProps(prev.xml, "w:pPr", { "w:keepNext": "<w:keepNext/>" }, PPR_ORDER);
    }
  };
  /** Blok berikutnya masih bagian tanda tangan ("Mengetahui," / tabel tanda tangan lain). */
  const continuesBelow = (at: number) => {
    const next = nextContent(at);
    if (!next) return false;
    if (next.kind === "table") return Boolean(opts.tables?.has(next.id));
    if (next.kind !== "paragraph") return false;
    const t = norm(next.text);
    return ROLE_RE.test(t) || DATE_RE.test(t);
  };

  let i = 0;
  while (i < model.body.length) {
    const cur = model.body[i];

    // 0. Tabel tanda tangan asli → dirapikan tanpa garis.
    if (cur.kind === "table" && opts.tables?.has(cur.id)) {
      cleanSignatureTable(cur, fmt, continuesBelow(i), opts.width);
      bindPrevious(i);
      rep.cleaned++;
      i++;
      continue;
    }

    // 1. Kumpulkan rangkaian paragraf yang mungkin blok tanda tangan.
    const lines: Line[] = [];
    let j = i;
    let gap = false;
    let lastLineAt = -1;
    while (j < model.body.length) {
      const n = model.body[j];
      if (n.kind !== "paragraph") break;
      const it = itemOf.get(n.id);
      if (!it || it.zone === "cover" || (it.role !== "body" && it.role !== "empty")) break;
      if (n.hasSectPr || n.hasPageBreak || n.hasDrawing || n.hasNumbering) break;
      if (isEmptyPara(n)) { if (lines.length) gap = true; j++; continue; }
      const parts = n.text.split(/\t+/).map((s) => s.trim()).filter(Boolean);
      const kinds = parts.map(classifyPart);
      if (!parts.length || parts.length > 4 || kinds.some((k) => k === null)) break;
      lines.push({ p: n, parts, kinds: kinds as K[], cls: alignClass(n, parts.length), gapBefore: gap });
      gap = false;
      lastLineAt = j;
      j++;
    }
    if (!lines.length) { i = Math.max(i + 1, j); continue; }

    // 2. Nama tanpa gelar / inisial: dikonfirmasi oleh konteks di sekitarnya.
    lines.forEach((l, li) => {
      l.kinds = l.kinds.map((k, c) => {
        if (k !== "plain") return k;
        const prev = lines[li - 1];
        const next = lines[li + 1];
        const prevK = prev?.kinds[Math.min(c, prev.kinds.length - 1)];
        const nextK = next?.kinds[Math.min(c, next.kinds.length - 1)];
        const prevT = prev?.parts[Math.min(c, prev.parts.length - 1)] ?? "";
        if (nextK === "id" || prevK === "materai" || prevK === "sigline" ||
          (prevK === "role" && (l.gapBefore || SELF_ROLE_RE.test(prevT)))) return "name";
        return k;
      });
    });

    const all = lines.flatMap((l) => l.kinds);
    const anchored = all.some((k) => k === "name" || k === "id" || k === "sigline");
    const context = all.some((k) => k === "role" || k === "date" || k === "materai");
    const firstAt = model.body.indexOf(lines[0].p);
    if (!anchored || !context || lines.length < 2) { i = Math.max(i + 1, lastLineAt + 1); continue; }

    // 3. Susun baris tabel.
    const N = Math.max(2, ...lines.map((l) => (l.cls === "multi" ? l.parts.length : 1)));
    type Row = { cells: Item[][]; span: boolean };
    const rows: Row[] = [];
    let curKey: Cls | null = null;
    for (const l of lines) {
      if (l.cls !== curKey) {
        rows.push({ cells: Array.from({ length: l.cls === "center" ? 1 : N }, () => []), span: l.cls === "center" });
        curKey = l.cls;
      }
      const row = rows[rows.length - 1];
      if (l.cls === "multi") {
        const cols = splitColumns(l.p.xml);
        // Baris berkolom lebih sedikit dari N yang diawali tab: rata ke kolom kanan.
        const shift = /^\t/.test(l.p.text) ? Math.max(0, N - cols.length) : 0;
        cols.slice(0, N).forEach((x, c) => row.cells[Math.min(N - 1, c + shift)].push({ xml: x, kind: l.kinds[Math.min(c, l.kinds.length - 1)] }));
      } else {
        const c = l.cls === "center" ? 0 : l.cls === "right" ? N - 1 : 0;
        row.cells[c].push({ xml: l.p.xml, kind: l.kinds[0] });
      }
    }

    // 4. Bangun XML tabel.
    const colW = Math.floor(opts.width / N);
    const afterLines: string[] = [];
    const afterTexts: string[] = [];
    let cellCount = 0;
    const prevNode = model.body[firstAt - 1];
    const needTopGap = Boolean(prevNode) && !(prevNode.kind === "paragraph" && isEmptyPara(prevNode));
    const lastAt = model.body.indexOf(lines[lines.length - 1].p);
    const keepAfter = continuesBelow(lastAt);
    const rowXml = rows.map((row, ri) => {
      const lastRow = ri === rows.length - 1;
      const hasSigner = row.cells.some((items) => items.some((x) => x.kind === "name" || x.kind === "id"));
      const rowGap = hasSigner ? SIGNER_ROW_GAP : ROW_GAP;
      const cells = row.cells.map((items) => {
        const sp = signatureSpacing(items.map((x) => x.kind));
        const paras = items.map((it, k) => {
          const s = sp.get(k) ?? {};
          const after = k === items.length - 1 ? (lastRow ? 0 : rowGap) : (s.after ?? 0);
          const before = ri === 0 && k === 0 && needTopGap ? Math.max(s.before ?? 0, BLOCK_GAP) : (s.before ?? 0);
          const x = fmtRuns(fmtPara(it.xml, before, after, !lastRow || keepAfter), fmt);
          const line = paragraphLine(x);
          if (line) { afterLines.push(line); afterTexts.push(line); }
          return x;
        });
        const span = row.span ? N : 1;
        cellCount++;
        return `<w:tc><w:tcPr><w:tcW w:w="${colW * span}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ""}<w:vAlign w:val="top"/></w:tcPr>${paras.length ? paras.join("") : EMPTY_P}</w:tc>`;
      });
      return `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cells.join("")}</w:tr>`;
    });
    const nil = (k: string) => `<w:${k} w:val="nil"/>`;
    const tblXml =
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/>` +
      `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map(nil).join("")}</w:tblBorders>` +
      `<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="57" w:type="dxa"/><w:right w:w="57" w:type="dxa"/></w:tblCellMar>` +
      `<w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>${`<w:gridCol w:w="${colW}"/>`.repeat(N)}</w:tblGrid>${rowXml.join("")}</w:tbl>`;

    // 5. Ganti paragraf blok dengan tabel.
    bindPrevious(firstAt);
    const replaced = model.body.slice(firstAt, lastAt + 1).filter((n): n is ParaNode => n.kind === "paragraph");
    const inserted = parseBlocks(
      (model.body[firstAt - 1]?.kind === "table" ? EMPTY_P : "") + tblXml + (lastAt + 1 >= model.body.length ? EMPTY_P : ""),
      1_000_000 + rep.blocks * 1000,
    );
    model.body.splice(firstAt, lastAt - firstAt + 1, ...inserted);

    rep.regions.push({
      beforeLines: replaced.map((p) => paragraphLine(p.xml)).filter(Boolean),
      afterLines,
      beforeText: replaced.map((p) => p.text.replace(/\t/g, " ")).join(" "),
      afterText: afterTexts.join(" "),
    });
    rep.blocks++;
    rep.added.tables++;
    rep.added.rows += rows.length;
    rep.added.cells += cellCount;

    i = firstAt + inserted.length;
  }

  return rep;
}