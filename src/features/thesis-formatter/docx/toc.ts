/**
 * Daftar isi / tabel / gambar / lampiran → field TOC Word yang bisa diklik
 * dan nomor halamannya dihitung Word.
 *
 * Entri lama TIDAK dihapus: dipakai sebagai hasil cache field (teks dokumen
 * tetap identik). Word mengganti isinya saat field diperbarui.
 * Sumber entri: heading 1–3 (\o \u) + penanda TC (\f) untuk judul bab dua
 * baris, keterangan tabel (t), gambar (g), dan lampiran (l).
 *
 * Daftar yang sudah berupa field TOC (termasuk yang dibungkus content control
 * "Table of Contents" bawaan Word) tidak ditumpuk field baru: pembungkusnya
 * dilepas dan instruksi field lamanya diganti.
 */
import { escapeXml } from "./xml";
import { flatten, type DocModel, type ParaNode, type TableNode } from "./document-model";
import type { Outline } from "../structure/outline";

export type TocReport = { tocs: number; entries: number; replaced: number; unwrapped: number };

const clean = (s: string) => escapeXml(s.replace(/\s+/g, " ").replace(/"/g, "'").trim());
const fld = (type: "begin" | "separate" | "end", dirty = false) =>
  `<w:r><w:fldChar w:fldCharType="${type}"${dirty ? ' w:dirty="true"' : ""}/></w:r>`;
const instr = (code: string) => `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>`;
const tc = (text: string, type: string) => fld("begin") + instr(`TC "${clean(text)}" \\f ${type} \\l 1`) + fld("end");

/** Paragraf memuat field TOC (bentuk fldChar maupun fldSimple). */
const TOC_INSTR_RE = /<w:instrText[^>]*>\s*TOC\b|<w:fldSimple\b[^>]*w:instr="\s*TOC\b/;

function append(xml: string, runs: string) {
  const i = xml.lastIndexOf("</w:p>");
  return i === -1 ? xml : xml.slice(0, i) + runs + xml.slice(i);
}
function prepend(xml: string, runs: string) {
  const e = xml.indexOf("</w:pPr>");
  if (e !== -1) return xml.slice(0, e + 8) + runs + xml.slice(e + 8);
  const o = xml.indexOf(">") + 1;
  return xml.slice(0, o) + runs + xml.slice(o);
}

/**
 * Content control yang membungkus field TOC dilepas: isinya jadi paragraf biasa
 * di body, supaya isi daftarnya bisa disusun ulang. Teks tidak berubah.
 */
export function unwrapTocControls(model: DocModel): number {
  let count = 0;
  const list = model.body;
  for (let i = 0; i < list.length; i++) {
    const b = list[i];
    if (b.kind !== "container") continue;
    if (!flatten(b.children).some((p) => p.kind === "paragraph" && TOC_INSTR_RE.test(p.xml))) continue;
    list.splice(i, 1, ...b.children);
    count++;
    i--; // periksa ulang anak yang baru masuk (content control bersarang)
  }
  return count;
}

/** Ganti instruksi field TOC pertama di paragraf dengan `code` dan tandai perlu diperbarui. */
function retargetToc(xml: string, code: string): string {
  if (/<w:fldSimple\b[^>]*w:instr="\s*TOC\b/.test(xml)) {
    return xml.replace(/(<w:fldSimple\b[^>]*w:instr=")[^"]*(")/, `$1 ${escapeXml(code)} $2`);
  }
  const runs = [...xml.matchAll(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g)];
  for (let k = 0; k < runs.length; k++) {
    if (!/w:fldCharType="begin"/.test(runs[k][0])) continue;
    const ins: number[] = [];
    let text = "";
    for (let s = k + 1; s < runs.length && !/w:fldCharType="(?:separate|end)"/.test(runs[s][0]); s++) {
      const m = /<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/.exec(runs[s][0]);
      if (m) { ins.push(s); text += m[1]; }
    }
    if (!ins.length || !/^\s*TOC\b/.test(text)) continue;

    let out = "";
    let last = 0;
    runs.forEach((m, idx) => {
      let r = m[0];
      if (idx === k) {
        if (/w:dirty=/.test(r)) return;
        r = r.replace(/w:fldCharType="begin"/, `w:fldCharType="begin" w:dirty="true"`);
      } else if (idx === ins[0]) {
        r = r.replace(/<w:instrText[^>]*>[\s\S]*?<\/w:instrText>/, `<w:instrText xml:space="preserve"> ${escapeXml(code)} </w:instrText>`);
      } else if (ins.includes(idx)) {
        r = r.replace(/(<w:instrText[^>]*>)[\s\S]*?(<\/w:instrText>)/, "$1$2");
      } else return;
      out += xml.slice(last, m.index) + r;
      last = m.index! + m[0].length;
    });
    return out + xml.slice(last);
  }
  return xml;
}

const TOC_CODE: [RegExp, string][] = [
  [/\bisi\b/i, 'TOC \\o "1-3" \\h \\z \\u \\f c'],
  [/tabel/i, "TOC \\h \\z \\f t"],
  [/gambar|grafik|bagan|ilustrasi/i, "TOC \\h \\z \\f g"],
  [/lampiran/i, "TOC \\h \\z \\f l"],
];
const NUMBERED_APPENDIX = /^lampiran\s*([0-9]+|[a-z])\b/i;

export function buildTocFields(
  model: DocModel,
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  opts: { upper: boolean; width: number },
): TocReport {
  const rep: TocReport = { tocs: 0, entries: 0, replaced: 0, unwrapped: 0 };
  rep.unwrapped = unwrapTocControls(model);

  const nextText = (i: number) => {
    for (let j = i + 1; j < nodes.length; j++) {
      const n = nodes[j];
      if (n.kind === "table") return -1;
      if (n.text.trim()) return j;
    }
    return -1;
  };

  // 1. Penanda TC
  nodes.forEach((n, i) => {
    if (n.kind !== "paragraph" || / TC "/.test(n.xml)) return;
    const it = outline.items[i];
    if (!it) return;
    let text: string | null = null;
    let type = "";
    if (it.role === "chapter_label" && it.pairedTitleNext) {
      const j = nextText(i);
      if (j !== -1) { text = `${n.text.trim()} ${(nodes[j] as ParaNode).text.trim()}`; type = "c"; }
    } else if (it.zone !== "front" && it.role === "table_caption") { text = n.text; type = "t"; }
    else if (it.zone !== "front" && it.role === "figure_caption") { text = n.text; type = "g"; }
    else if (it.role === "appendix_heading" && NUMBERED_APPENDIX.test(n.text.trim())) { text = n.text; type = "l"; }
    if (!text) return;
    if (type === "c" && opts.upper) text = text.toUpperCase();
    n.xml = append(n.xml, tc(text, type));
    n.hasField = true;
    rep.entries++;
  });

  // 2. Daftar lama → field TOC
  nodes.forEach((n, i) => {
    if (n.kind !== "paragraph" || outline.items[i]?.role !== "list_heading") return;
    const code = TOC_CODE.find(([re]) => re.test(n.text))?.[1];
    if (!code) return;

    const entries: ParaNode[] = [];
    let existing: ParaNode | null = null;
    for (let j = i + 1; j < nodes.length; j++) {
      const m = nodes[j];
      const r = outline.items[j]?.role;
      if (m.kind !== "paragraph") break;
      if (!existing && TOC_INSTR_RE.test(m.xml)) existing = m;
      if (r === "list_entry") entries.push(m);
      else if (r !== "empty") break;
    }

    // Sudah ada field TOC (mis. "No table of contents entries found." atau daftar isi
    // otomatis lama): cukup ganti instruksinya, jangan tambah field kedua.
    if (existing) {
      const x = retargetToc(existing.xml, code);
      if (x !== existing.xml) { existing.xml = x; existing.hasField = true; rep.replaced++; }
      rep.tocs++;
      return;
    }

    const begin = fld("begin", true) + instr(code) + fld("separate");
    if (entries.length) {
      entries[0].xml = prepend(entries[0].xml, begin);
      const last = entries[entries.length - 1];
      last.xml = append(last.xml, fld("end"));
      entries[0].hasField = last.hasField = true;
    } else {
      const idx = model.body.indexOf(n);
      if (idx === -1) return;
      const xml = `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="${opts.width}"/></w:tabs></w:pPr>${begin}${fld("end")}</w:p>`;
      model.body.splice(idx + 1, 0, {
        kind: "paragraph", id: -(i + 1), xml, text: "", hasDrawing: false, hasPageBreak: false,
        hadPageBreakBefore: false, hasSectPr: false, hasNumbering: false, hasField: true,
        styleId: null, align: null, tabCount: 0, allBold: false,
      });
    }
    rep.tocs++;
  });

  return rep;
}