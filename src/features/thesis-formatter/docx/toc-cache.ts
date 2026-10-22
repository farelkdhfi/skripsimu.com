/**
 * Isi tampilan (cached result) field DAFTAR TABEL / DAFTAR GAMBAR.
 *
 * Field TOC sudah ditandai dirty, tetapi penampil yang tidak memperbarui field
 * (Google Docs, WPS, LibreOffice, pratinjau — atau Word bila pembaruan ditolak)
 * tetap menampilkan isi tersimpan, yang sebelumnya masih daftar lama dokumen
 * asli. Di sini isinya disusun ulang dari entri TC yang benar-benar ada.
 */
import { parseBlocks, type DocModel, type ParaNode } from "./document-model";
import { escapeXml } from "./xml";
import { paragraphLine, type VerifyRegion } from "./verify";
import { stripFields } from "./tables";

type Flag = "t" | "g";
type Body = DocModel["body"];
type Block = { i0: number; i1: number; open: string };

const decodeXml = (s: string) =>
  s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const runsOf = (xml: string) => [...xml.matchAll(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g)];
const instrOf = (xml: string) =>
  [...xml.matchAll(/<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/g)].map((m) => decodeXml(m[1])).join("");
const textOf = (xml: string) =>
  [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => decodeXml(m[1])).join("");
const pPrOf = (xml: string) => /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? "";
const keyOf = (t: string) =>
  t.replace(/^\s*(?:tabel|table|gambar|figure)\s*\.?\s*(?:(?:[A-Za-z]{1,2}[.\-]?)?\d[\d.\-]*)?\s*[.:]?\s*/i, "")
    .toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

function collectTc(body: Body, flag: Flag): string[] {
  const re = new RegExp(String.raw`\bTC\s+"([^"]*)"[^"]*?\\f\s*${flag}(?![a-z])`, "g");
  const out: string[] = [];
  for (const n of body) {
    if (n.kind !== "paragraph" || !/instr/.test(n.xml)) continue;
    const ins = instrOf(n.xml) + " " + [...n.xml.matchAll(/w:instr="([^"]*)"/g)].map((m) => decodeXml(m[1])).join(" ");
    for (const m of ins.matchAll(re)) out.push(m[1].replace(/\s+/g, " ").trim());
  }
  return out;
}

function locate(body: Body, flag: Flag): Block | null {
  const tocRe = new RegExp(String.raw`^\s*TOC\b[\s\S]*\\f\s*${flag}(?![a-z])`);
  for (let i = 0; i < body.length; i++) {
    const n = body[i];
    if (n.kind !== "paragraph" || !n.xml.includes("TOC")) continue;
    const runs = runsOf(n.xml);
    for (let k = 0; k < runs.length; k++) {
      if (!/w:fldCharType="begin"/.test(runs[k][0])) continue;
      let ins = "";
      let s = k + 1;
      for (; s < runs.length && !/w:fldCharType="(?:separate|end)"/.test(runs[s][0]); s++) ins += instrOf(runs[s][0]);
      if (s >= runs.length || !/separate/.test(runs[s][0]) || !tocRe.test(ins)) continue;
      if (textOf(n.xml.slice(0, runs[k].index)).trim()) return null;
      const openEnd = runs[s].index! + runs[s][0].length;
      const open = n.xml.slice(runs[k].index!, openEnd);

      let depth = 1;
      const scan = (xml: string, from: number) => {
        for (const m of xml.slice(from).matchAll(/w:fldCharType="(begin|end)"/g)) {
          depth += m[1] === "begin" ? 1 : -1;
          if (depth === 0) return from + m.index!;
        }
        return -1;
      };
      let at = scan(n.xml, openEnd);
      let j = i;
      while (at === -1 && ++j < body.length) {
        const b = body[j];
        if (b.kind !== "paragraph") return null;
        at = scan(b.xml, 0);
      }
      if (at === -1) return null;
      const tail = (body[j] as ParaNode).xml.slice(at);
      if (textOf(tail.slice(tail.indexOf("</w:r>") + 6)).trim()) return null;
      return { i0: i, i1: j, open };
    }
  }
  return null;
}

function oldEntries(body: Body, b: Block) {
  const out: { label: string; key: string; page: string; used: boolean }[] = [];
  for (let i = b.i0; i <= b.i1; i++) {
    const t = (body[i] as ParaNode).text.replace(/\s+$/, "");
    const m = /^(.*?)(?:\t|\.{2,}|…|\s{2,})+\s*([0-9]+|[ivxlcdm]+)$/i.exec(t) ?? /^(.*?\D)(\d{1,3})$/.exec(t);
    if (m) out.push({ label: m[1].trim(), key: keyOf(m[1]), page: m[2], used: false });
    else if (t.trim()) out.push({ label: t.trim(), key: keyOf(t), page: "", used: false });
  }
  return out;
}

export function refreshListCaches(model: DocModel, opts: { dropTc?: Set<number> } = {}) {
  const regions: VerifyRegion[] = [];
  const counts = { tables: 0, figures: 0 };
  const body = model.body;

  // TC di keterangan lampiran tidak ikut Daftar Tabel (lampiran punya daftarnya sendiri).
  if (opts.dropTc?.size) {
    for (const n of body) if (n.kind === "paragraph" && opts.dropTc.has(n.id)) n.xml = stripFields(n.xml, /^\s*TC\b/i, false);
  }

    let nextId = 1_000_000 + body.reduce(
    (m, n) => (n.kind === "paragraph" || n.kind === "table" ? Math.max(m, n.id) : m), 0);
  for (const flag of ["t", "g"] as const) {
    const entries = collectTc(body, flag);
    const b = entries.length ? locate(body, flag) : null;
    if (!b) continue;

    const olds = oldEntries(body, b);
    if (olds.length === entries.length && olds.every((o, i) => o.label === entries[i])) continue;

    const items = entries.map((text) => {
      const k = keyOf(text);
      let best: (typeof olds)[number] | null = null;
      for (const o of olds) {
        if (o.used || o.key.length < 5 || k.length < 5) continue;
        if (!(k.startsWith(o.key) || o.key.startsWith(k))) continue;
        if (!best || o.key.length > best.key.length) best = o;
      }
      if (best) best.used = true;
      return { text, page: best?.page ?? "" };
    });

    const first = body[b.i0] as ParaNode;
    const last = body[b.i1] as ParaNode;
    const noSect = (p: string) => p.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/, "");
    const firstPPr = noSect(pPrOf(first.xml));
    const midPPr = firstPPr.replace(/<w:pageBreakBefore(?:\s[^>]*)?\/>/, "");
    const lastPPr = /<w:sectPr\b/.test(pPrOf(last.xml)) ? pPrOf(last.xml) : midPPr;
    const rPr = (() => {
      for (const m of runsOf(first.xml + last.xml)) {
        if (/<w:t[\s>]/.test(m[0]) && !/fldChar|instrText/.test(m[0]))
          return (/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(m[0])?.[0] ?? "")
            .replace(/<w:rStyle\b[^>]*\/>/, "").replace(/<w:color\b[^>]*\/>/, "").replace(/<w:u\b[^>]*\/>/, "");
      }
      return "";
    })();
    const run = (inner: string) => `<w:r>${rPr}${inner}</w:r>`;

    const xml = items.map((e, idx) => {
      const isFirst = idx === 0;
      const isLast = idx === items.length - 1;
      const pPr = isLast ? (isFirst && !/<w:sectPr\b/.test(lastPPr) ? firstPPr : lastPPr) : isFirst ? firstPPr : midPPr;
      let inner = run(`<w:t xml:space="preserve">${escapeXml(e.text)}</w:t>`) + run("<w:tab/>") + (e.page ? run(`<w:t>${e.page}</w:t>`) : "");
      if (isFirst) inner = b.open + inner;
      if (isLast) inner += `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
      return `<w:p>${pPr}${inner}</w:p>`;
    }).join("");

    const fresh = parseBlocks(xml, nextId) as ParaNode[];
    nextId += fresh.length + 10;
    const removed = body.slice(b.i0, b.i1 + 1) as ParaNode[];
    regions.push({
      beforeLines: removed.map((p) => paragraphLine(p.xml)).filter(Boolean),
      afterLines: fresh.map((p) => paragraphLine(p.xml)).filter(Boolean),
      beforeText: "",
      afterText: "",
    });
    body.splice(b.i0, b.i1 - b.i0 + 1, ...fresh);
    if (flag === "t") counts.tables = items.length; else counts.figures = items.length;
  }
  return { ...counts, regions };
}