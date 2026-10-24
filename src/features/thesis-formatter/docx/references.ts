/**
 * Daftar Pustaka.
 *  - APA/Harvard: urut abjad nama penulis, nomor "[1]"/"1." dibuang,
 *    nama penulis kapital semua dikembalikan ke huruf biasa.
 *  - IEEE/Vancouver: urutan mengikuti nomor sitasi → tidak diubah.
 *  - Entri yang gayanya tidak sesuai pilihan dilaporkan (isi tidak diubah).
 */
import type { DocModel, ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { isHeadingRole } from "../structure/roles";
import { paragraphLine, type VerifyRegion } from "./verify";
import { leadingRaw, removeLeading } from "./lists";
import { retext } from "./tables";

export type CitationStyle = "APA" | "IEEE" | "Harvard" | "Vancouver";
type Kind = "numbered" | "apa" | "harvard" | "vancouver" | "other";

const YEAR_RE = /\b(?:1[89]|20)\d{2}[a-z]?\b|n\.d\.|t\.t\./i;
const NUM_PREFIX_RE = /^[ \t\u00a0]*(?:\[\d{1,3}\]|\d{1,3}[.)])[ \t\u00a0]+/;
const NOT_ENTRY_RE = /^(catatan|keterangan|note|sumber)\b/i;

function kindOf(t: string): Kind {
  if (NUM_PREFIX_RE.test(t)) return "numbered";
  const head = t.slice(0, 220);
  if (/\((?:(?:1[89]|20)\d{2}[a-z]?|n\.d\.)\)/.test(head)) return "apa";
  if (/^[A-Z][A-Za-z'’-]+ [A-Z]{1,3}[,.]/.test(t)) return "vancouver";
  if (/[,.]\s(?:1[89]|20)\d{2}[a-z]?[.,]/.test(head)) return "harvard";
  return "other";
}

const fits = (k: Kind, s: CitationStyle) =>
  s === "APA" ? k === "apa"
  : s === "Harvard" ? k === "harvard" || k === "apa"
  : s === "IEEE" ? k === "numbered"
  : k === "numbered" || k === "vancouver";

function sortKey(p: ParaNode): string {
  let t = p.text.replace(NUM_PREFIX_RE, "").replace(/^[\s"“'‘(\[]+/, "");
  const ini = /^((?:[A-Z]\.\s?-?){1,4})\s*([A-Z][^\s,]+)/.exec(t);   // "J. Nielsen" → Nielsen
  if (ini) t = `${ini[2]} ${t}`;
  return t.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

/** "HAIR, J.F., BLACK, W.C. dan ANDERSON, R.E. (2019)" → "Hair, J.F., Black, W.C. dan Anderson, R.E. (2019)". */
function fixAuthorCaps(full: string): string {
  const m = /^(.*?)(?=\(\s*(?:1[89]|20)\d{2}|[.,]\s*(?:1[89]|20)\d{2}\b)/.exec(full);
  const seg = m?.[1] ?? "";
  if (!seg || seg.length > 250) return full;
  const check = seg.replace(/\b(dan|and|et|al|dkk|&)\b\.?/gi, "");
  if (/[a-z]/.test(check) || !/[A-Z]{2,},\s*[A-Z]\./.test(check)) return full;
  const fixed = seg.replace(/\b([A-Z])([A-Z'’-]+)\b(?!\.)/g, (_w, a: string, b: string) => a + b.toLowerCase());
  return fixed + full.slice(seg.length);
}

export type ReferenceReport = {
  entries: number;
  sorted: boolean;
  numbersRemoved: number;
  capsFixed: number;
  mismatched: string[];
  regions: VerifyRegion[];
};

export function arrangeReferences(
  model: DocModel,
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  style: CitationStyle,
): ReferenceReport {
  const rep: ReferenceReport = { entries: 0, sorted: false, numbersRemoved: 0, capsFixed: 0, mismatched: [], regions: [] };
  const start = nodes.findIndex((_n, i) => outline.items[i]?.role === "reference_heading");
  if (start === -1) return rep;

  const entries: ParaNode[] = [];
  for (let i = start + 1; i < nodes.length; i++) {
    const it = outline.items[i];
    if (it && isHeadingRole(it.role)) break;
    const n = nodes[i];
    if (n.kind !== "paragraph") continue;
    const t = n.text.trim();
    if (t.length < 25 || !YEAR_RE.test(t) || NOT_ENTRY_RE.test(t)) continue;
    entries.push(n);
  }
  rep.entries = entries.length;
  if (entries.length < 2) return rep;

  const beforeLines = entries.map((p) => paragraphLine(p.xml)).filter(Boolean);
  for (const p of entries) {
    if (!fits(kindOf(p.text.trim()), style)) rep.mismatched.push(p.text.trim().slice(0, 70));
  }

  if (style === "APA" || style === "Harvard") {
    for (const p of entries) {
      const m = NUM_PREFIX_RE.exec(leadingRaw(p.xml, 12));
      if (m) {
        p.xml = removeLeading(p.xml, m[0].length);
        p.text = p.text.replace(NUM_PREFIX_RE, "");
        rep.numbersRemoved++;
      }
      const xml = retext(p.xml, fixAuthorCaps);
      if (xml !== p.xml) { p.xml = xml; p.text = fixAuthorCaps(p.text); rep.capsFixed++; }
    }

    const pos = entries.map((p) => model.body.indexOf(p));
    if (pos.every((k) => k >= 0)) {
      const sorted = [...entries].sort((a, b) => sortKey(a).localeCompare(sortKey(b), "id"));
      if (sorted.some((p, k) => p !== entries[k])) {
        [...pos].sort((a, b) => a - b).forEach((k, j) => { model.body[k] = sorted[j]; });
        rep.sorted = true;
      }
    }
  }

  if (rep.sorted || rep.numbersRemoved || rep.capsFixed) {
    const ordered = model.body.filter((n): n is ParaNode => n.kind === "paragraph" && entries.includes(n as ParaNode));
    const text = ordered.map((p) => p.text).join(" ");
    rep.regions.push({
      beforeLines,
      afterLines: ordered.map((p) => paragraphLine(p.xml)).filter(Boolean),
      beforeText: text,
      afterText: text,
    });
  }
  return rep;
}