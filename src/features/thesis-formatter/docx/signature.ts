/**
 * Blok tanda tangan → tabel tanpa border dengan ruang tanda tangan yang cukup.
 *
 * Deteksi: rangkaian paragraf pendek berurutan yang berisi tanggal, jabatan,
 * materai, nama (bergelar/terkonfirmasi konteks), dan nomor identitas.
 * Hanya diubah bila memuat NAMA/NOMOR IDENTITAS dan JABATAN/TANGGAL/MATERAI.
 *
 * Tata letak:
 *  - baris bertab (beberapa penanda tangan sejajar) → satu kolom per penanda tangan
 *  - baris rata tengah → satu sel selebar tabel
 *  - baris rata kanan / menjorok kanan → kolom kanan
 *  - ruang tanda tangan ±2 cm sebelum nama (atau mengapit materai)
 */
import { childEl, innerXml, openTag, splitElements, withProps, PPR_ORDER } from "./xml";
import { parseBlocks, type DocModel, type ParaNode, type TableNode } from "./document-model";
import type { Outline, OutlineItem } from "../structure/outline";
import { paragraphLine } from "./verify";

export type SigRegion = { beforeLines: string[]; afterLines: string[]; beforeText: string; afterText: string };
export type SigReport = { blocks: number; regions: SigRegion[]; added: { tables: number; rows: number; cells: number } };

type K = "date" | "role" | "materai" | "id" | "name" | "plain";
type Cls = "multi" | "left" | "center" | "right";
type Line = { p: ParaNode; parts: string[]; kinds: K[]; cls: Cls; gapBefore: boolean };
type Item = { xml: string; kind: K };

const DATE_RE = /^[A-Z][A-Za-z.'\- ]{1,40},\s*(\d{1,2}\s+)?[A-Za-z]{3,12}\s+\d{4}\.?$/;
const ROLE_RE = /^(mengetahui|menyetujui|disetujui|disahkan|pembimbing|penguji|ketua|dekan|wakil|rektor|direktur|kepala|sekretaris|koordinator|kaprodi|yang\s+menyatakan|yang\s+membuat|penulis|peneliti|mahasiswa|dosen|pembina|atas\s+nama|a\.n\.)/i;
const SELF_ROLE_RE = /(yang\s+menyatakan|yang\s+membuat|penulis|peneliti|mahasiswa)/i;
const MATERAI_RE = /^\[?\s*(materai|meterai)/i;
const ID_RE = /^(NIP|NIDN|NIDK|NIK|NIM|NPM|NRP|NISN|NUPTK)\b/i;
const DEGREE_RE = /(\b(Prof|Dr|Drs|Dra|Ir|H|Hj)\.\s)|(,\s*(S|M|Dr|Ph|Sp|Ak|Apt)\.?\s?[A-Za-z]{0,6}\.?)|(\bPh\.?D\b)/;
const PLAIN_RE = /^[A-Z][A-Za-z.'\u2019\-]*( [A-Z][A-Za-z.'\u2019\-]*){1,5}$/;

const SIG_SPACE = 1300;        // ±3 cm ruang tanda tangan
const MATERAI_PAD = 600;       // jarak di atas & bawah tulisan materai
const ROW_GAP = 240;           // setelah baris tanpa penanda tangan (mis. tanggal)
const SIGNER_ROW_GAP = 567;    // ±1 cm setelah baris penanda tangan, sebelum blok berikutnya
const BLOCK_GAP = 567;         // ±1 cm antara teks isi dan blok tanda tangan

const EMPTY_P = `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>`;

function classifyPart(t: string): K | null {
  if (!t || t.length > 90) return null;
  if (ID_RE.test(t)) return "id";
  if (MATERAI_RE.test(t)) return "materai";
  if (DATE_RE.test(t)) return "date";
  if (ROLE_RE.test(t) && t.length <= 80) return "role";
  if (t.includes(":")) return null;
  if (DEGREE_RE.test(t) && t.length <= 80) return "name";
  if (PLAIN_RE.test(t) && t.length <= 60) return "plain";
  return null;
}

function alignClass(p: ParaNode, parts: number): Cls {
  if (parts > 1) return "multi";
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
    "w:pageBreakBefore": null,
    "w:numPr": null,
    "w:tabs": null,
    "w:spacing": `<w:spacing w:before="${before}" w:after="${after}" w:line="240" w:lineRule="auto"/>`,
    "w:ind": `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`,
    "w:jc": `<w:jc w:val="center"/>`,
  }, PPR_ORDER);
}

/** Ruang tanda tangan di dalam satu sel. */
function signatureSpacing(items: Item[]): Map<number, { before?: number; after?: number }> {
  const out = new Map<number, { before?: number; after?: number }>();
  let nameIdx = items.findIndex((x) => x.kind === "id") - 1;
  if (nameIdx < 0) nameIdx = items.findIndex((x) => x.kind === "name");
  if (nameIdx < 0) {
    const m = items.findIndex((x) => x.kind === "materai");
    if (m !== -1 && m + 1 < items.length) nameIdx = m + 1;
  }
  if (nameIdx <= 0) return out;
  if (items[nameIdx - 1].kind === "materai") out.set(nameIdx - 1, { before: MATERAI_PAD, after: MATERAI_PAD });
  else out.set(nameIdx, { before: SIG_SPACE });
  return out;
}

const isEmptyPara = (p: ParaNode) =>
  !p.text.trim() && !p.hasDrawing && !p.hasSectPr && !p.hasPageBreak && !p.hasField;

export function buildSignatureTables(
  model: DocModel,
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  opts: { width: number },
): SigReport {
  const rep: SigReport = { blocks: 0, regions: [], added: { tables: 0, rows: 0, cells: 0 } };
  const itemOf = new Map<number, OutlineItem>();
  nodes.forEach((n, i) => { if (outline.items[i]) itemOf.set(n.id, outline.items[i]); });

  let i = 0;
  while (i < model.body.length) {
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
      if (!parts.length || kinds.some((k) => k === null)) break;
      lines.push({ p: n, parts, kinds: kinds as K[], cls: alignClass(n, parts.length), gapBefore: gap });
      gap = false;
      lastLineAt = j;
      j++;
    }
    if (!lines.length) { i = Math.max(i + 1, j); continue; }

    // 2. Nama tanpa gelar: dikonfirmasi oleh konteks di sekitarnya.
    lines.forEach((l, li) => {
      l.kinds = l.kinds.map((k, c) => {
        if (k !== "plain") return k;
        const prev = lines[li - 1];
        const next = lines[li + 1];
        const prevK = prev?.kinds[Math.min(c, prev.kinds.length - 1)];
        const nextK = next?.kinds[Math.min(c, next.kinds.length - 1)];
        const prevT = prev?.parts[Math.min(c, prev.parts.length - 1)] ?? "";
        if (nextK === "id" || prevK === "materai" || (prevK === "role" && (l.gapBefore || SELF_ROLE_RE.test(prevT)))) return "name";
        return k;
      });
    });

    const all = lines.flatMap((l) => l.kinds);
    const anchored = all.some((k) => k === "name" || k === "id");
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
        cols.slice(0, N).forEach((x, c) => row.cells[c].push({ xml: x, kind: l.kinds[Math.min(c, l.kinds.length - 1)] }));
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
    // Beri jarak dari teks di atasnya, kecuali sudah ada baris kosong di sana.
    const prevNode = model.body[firstAt - 1];
    const needTopGap = Boolean(prevNode) && !(prevNode.kind === "paragraph" && isEmptyPara(prevNode));
    const rowXml = rows.map((row, ri) => {
      const lastRow = ri === rows.length - 1;
      // Baris yang berisi penanda tangan diberi jarak lebih lega ke blok di bawahnya.
      const hasSigner = row.cells.some((items) => items.some((x) => x.kind === "name" || x.kind === "id"));
      const rowGap = hasSigner ? SIGNER_ROW_GAP : ROW_GAP;
      const cells = row.cells.map((items) => {
        const sp = signatureSpacing(items);
        const paras = items.map((it, k) => {
          const s = sp.get(k) ?? {};
          const after = k === items.length - 1 ? (lastRow ? 0 : rowGap) : (s.after ?? 0);
          const before = ri === 0 && k === 0 && needTopGap ? Math.max(s.before ?? 0, BLOCK_GAP) : (s.before ?? 0);
          const x = fmtPara(it.xml, before, after, !lastRow);
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
    const lastAt = model.body.indexOf(lines[lines.length - 1].p);
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