/**
 * Merapikan judul bab/subbab:
 *  - nomor subbab yang salah ketik (C. di antara 1.2 dan 1.4 → 1.3), hanya bila
 *    nomor tetangga memastikan slotnya; level ikut dikoreksi (2.1.2 = Heading 3)
 *  - nomor ganda/mundur ("2.1.3" dua kali → 2.1.4) bila nomor sesudahnya tidak bertentangan
 *  - tanda setelah nomor dibuang ("1.4." / "1.4)" → "1.4")
 *  - awalan bab angka → romawi ("BAB 2" → "BAB II")
 *  - judul bab satu baris ("BAB II. TINJAUAN PUSTAKA") → dua baris ("BAB II" ↵ "TINJAUAN PUSTAKA")
 *  - judul subbab Title Case (akronim dipertahankan)
 */
import { escapeXml } from "./xml";
import { formatParagraph, type ApplyReport } from "./apply-format";
import { leadingRaw, removeLeading } from "./lists";
import { retext, titleCaseText } from "./tables";
import { paragraphLine, type VerifyRegion } from "./verify";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import type { FormatRules } from "../rules/format-rules";

export type HeadingReport = {
  renumbered: number; retitled: number; relevelled: number;
  /** Judul bab satu baris yang dipecah jadi dua baris. */
  chaptersSplit: number;
  regions: VerifyRegion[];
};

const SEC_PREFIX_RE = /^[ \t\u00a0]*(\d+(?:\.\d+){1,3}|[A-Z])[.)]?[ \t\u00a0]+/;
const SEC_PREFIX_TEXT_RE = /^\s*(\d+(?:\.\d+){1,3}|[A-Z])[.)]?\s+/;
const CHAPTER_RAW_RE = /^[ \t\u00a0]*bab[ \t\u00a0]+([ivxlcdm]+|\d+)\b/i;
/** "BAB II. ", "BAB 3 : ", "Bab IV – " (label + pemisah sebelum judul). */
const CHAPTER_LINE_RE = /^[ \t\u00a0]*bab[ \t\u00a0]+([ivxlcdm]+|\d+)[ \t\u00a0]*[.:\-\u2013\u2014]?[ \t\u00a0]*/i;
const CHAPTER_TEXT_RE = /^\s*bab\s+([ivxlcdm]+|\d+)\s*[.:\-\u2013\u2014]?\s*/i;
const SOFT_BREAK_RE = /<w:br(?![^>]*w:type="page")[^>]*\/>|<w:cr\s*\/>/;

function toRoman(n: number): string {
  const map: [number, string][] = [[10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let out = "";
  for (const [v, s] of map) while (n >= v) { out += s; n -= v; }
  return out;
}
function fromRoman(s: string): number {
  const v: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
  const a = s.toLowerCase();
  let t = 0;
  for (let k = 0; k < a.length; k++) { const c = v[a[k]]; const n = v[a[k + 1]] ?? 0; t += c < n ? -c : c; }
  return t;
}

/** Ganti awalan paragraf (oldLen karakter, termasuk tab) dengan `prefix`. */
function setPrefix(p: ParaNode, oldLen: number, prefix: string) {
  const xml = removeLeading(p.xml, oldLen);
  p.xml = xml.replace(/<w:t(\s[^>]*)?>/, `<w:t xml:space="preserve">${escapeXml(prefix)}`);
}

/** Buang awalan (oldLen karakter), sisipkan run "label" + ganti baris di awal paragraf. */
function setLabelLine(p: ParaNode, oldLen: number, label: string) {
  let xml = removeLeading(p.xml, oldLen);
  const rPr = /<w:r(?:\s[^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/.exec(xml)?.[1] ?? "";
  const e = xml.indexOf("</w:pPr>");
  const at = e !== -1 ? e + 8 : xml.indexOf(">") + 1;
  xml = xml.slice(0, at) + `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(label)}</w:t><w:br/></w:r>` + xml.slice(at);
  p.xml = xml;
}

type Entry = { i: number; p: ParaNode; level: 2 | 3; nums: number[] | null; lettered: boolean; dotted: boolean; changed: boolean };

export function normalizeHeadings(
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  rules: FormatRules,
  rep: ApplyReport,
  opts: { titleCase: boolean },
): HeadingReport {
  const out: HeadingReport = { renumbered: 0, retitled: 0, relevelled: 0, chaptersSplit: 0, regions: [] };
  let chapterNo = 0;
  let entries: Entry[] = [];

  const resolve = () => {
    if (!entries.length) return;
    const numeric = entries.filter((e) => e.nums);
    const c = chapterNo || numeric[0]?.nums?.[0] || 0;

    // Awalan bab yang salah (mis. "1.3" di dalam BAB II), bila mayoritas sudah benar.
    if (c && numeric.filter((e) => e.nums![0] === c).length * 2 >= numeric.length) {
      for (const e of numeric) if (e.nums![0] !== c) { e.nums = [c, ...e.nums!.slice(1)]; e.changed = true; }
    }

    // Rangkaian judul berhuruf (A., B., C.) di antara judul bernomor.
    for (let k = 0; k < entries.length;) {
      if (!entries[k].lettered) { k++; continue; }
      let e2 = k;
      while (e2 < entries.length && entries[e2].lettered) e2++;
      const run = entries.slice(k, e2);
      const m = run.length;
      const P = k > 0 ? entries[k - 1].nums : null;
      const N = e2 < entries.length ? entries[e2].nums : null;
      const seq = (base: number[], start: number) => run.map((_, j) => [...base, start + j]);
      let assign: number[][] | null = null;
      if (c && P && N) {
        if (P.length === 2 && N.length === 2 && N[1] - P[1] === m + 1) assign = seq([c], P[1] + 1);
        else if (N.length === 3 && N[1] === P[1]) {
          const base = P.length === 3 ? P[2] : 0;
          if (N[2] - base === m + 1) assign = seq([c, P[1]], base + 1);
        } else if (P.length === 3 && N.length === 2 && N[1] === P[1] + 1) assign = seq([c, P[1]], P[2] + 1);
      } else if (c && P && !N) {
        assign = P.length === 2 ? seq([c], P[1] + 1) : seq([c, P[1]], P[2] + 1);
      } else if (c && !P && N) {
        if (N.length === 2 && N[1] === m + 1) assign = seq([c], 1);
        else if (N.length === 3 && N[2] === m + 1) assign = seq([c, N[1]], 1);
      }
      if (assign) run.forEach((e, j) => {
        e.nums = assign![j];
        e.level = e.nums.length >= 3 ? 3 : 2;
        e.lettered = false;
        e.changed = true;
      });
      k = e2;
    }

    // Nomor ganda atau mundur di bawah induk yang sama ("2.1.3" dua kali) → nomor
    // berikutnya, hanya bila saudara sesudahnya tidak bertentangan (2.1.5 → slot 2.1.4 kosong).
    const lastOf = new Map<string, number>();
    entries.forEach((e, idx) => {
      if (!e.nums) return;
      const parent = e.nums.slice(0, -1).join(".");
      const own = e.nums[e.nums.length - 1];
      const last = lastOf.get(parent);
      if (last !== undefined && own <= last) {
        const want = last + 1;
        const next = entries.slice(idx + 1).find((x) =>
          x.nums && x.nums.length === e.nums!.length && x.nums.slice(0, -1).join(".") === parent);
        const nextOwn = next?.nums?.[next.nums.length - 1];
        if (nextOwn === undefined || nextOwn > want) {
          e.nums = [...e.nums.slice(0, -1), want];
          e.changed = true;
        }
      }
      lastOf.set(parent, e.nums[e.nums.length - 1]);
    });

    for (const e of entries) {
      const p = e.p;
      const before = { line: paragraphLine(p.xml), rest: p.text.replace(SEC_PREFIX_TEXT_RE, "") };
      const m = SEC_PREFIX_RE.exec(leadingRaw(p.xml, 40));
      const fixPrefix = Boolean(m && e.nums && (e.changed || e.dotted));
      if (fixPrefix) setPrefix(p, m![0].length, `${e.nums!.join(".")} `);
      if (opts.titleCase) {
        const xml = retext(p.xml, titleCaseText);
        if (xml !== p.xml) { p.xml = xml; out.retitled++; }
      }
      const rest = titleCaseText(before.rest);
      p.text = fixPrefix ? `${e.nums!.join(".")} ${rest}` : (opts.titleCase ? titleCaseText(p.text) : p.text);

      if (fixPrefix) {
        out.renumbered++;
        out.regions.push({ beforeLines: [before.line].filter(Boolean), afterLines: [paragraphLine(p.xml)].filter(Boolean), beforeText: before.rest, afterText: rest });
      }

      // Level ikut dikoreksi (mis. "A." yang ternyata anak subbab).
      const it = outline.items[e.i];
      const want = e.level === 3 ? "subsection" : "section";
      if (it.role !== want) {
        it.role = want;
        formatParagraph(p, rules.paragraph(want, "main", it.hint), rules.run(want, "main", it.hint), null, rep);
        out.relevelled++;
      }
    }
    entries = [];
  };

  nodes.forEach((n, i) => {
    const it = outline.items[i];
    if (!it || it.zone !== "main" || n.kind !== "paragraph") return;

    if (it.role === "chapter_label" || (it.role === "chapter" && !it.continuesChapterLabel)) {
      resolve();
      const raw = leadingRaw(n.xml, 60);
      const m = CHAPTER_RAW_RE.exec(raw);
      if (!m) { chapterNo++; return; }
      const arabic = /^\d+$/.test(m[1]);
      chapterNo = arabic ? Number(m[1]) : fromRoman(m[1]) || chapterNo + 1;
      const label = `BAB ${toRoman(chapterNo)}`;
      const title = n.text.replace(CHAPTER_TEXT_RE, "").trim();
      // Judul bab satu baris → "BAB II" ↵ "TINJAUAN PUSTAKA" (seragam dengan bab dua baris).
      const oneLine = it.role === "chapter" && !it.pairedTitleNext && Boolean(title) && !SOFT_BREAK_RE.test(n.xml);
      if (!arabic && !oneLine) return;

      const beforeLine = paragraphLine(n.xml);
      if (oneLine) {
        const lm = CHAPTER_LINE_RE.exec(raw);
        if (!lm) return;
        setLabelLine(n, lm[0].length, label);
        n.text = `${label} ${title}`;
        out.chaptersSplit++;
      } else {
        const beforeRest = n.text.replace(/^\s*bab\s+\d+/i, "");
        setPrefix(n, m[0].length, label);
        n.text = `${label}${beforeRest}`;
        out.renumbered++;
      }
      out.regions.push({ beforeLines: [beforeLine].filter(Boolean), afterLines: [paragraphLine(n.xml)].filter(Boolean), beforeText: title, afterText: title });
      return;
    }

    if (it.role !== "section" && it.role !== "subsection") return;
    const m = SEC_PREFIX_RE.exec(leadingRaw(n.xml, 40));
    if (!m) return;
    const tok = m[1];
    const lettered = /^[A-Z]$/.test(tok);
    const nums = lettered ? null : tok.split(".").map(Number);
    entries.push({
      i, p: n, lettered, nums, changed: false,
      level: nums ? (nums.length >= 3 ? 3 : 2) : it.role === "subsection" ? 3 : 2,
      dotted: /[.)][ \t\u00a0]+$/.test(m[0]),
    });
  });
  resolve();
  return out;
}