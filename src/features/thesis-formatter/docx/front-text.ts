/**
 * Baris kata kunci abstrak: hanya label yang tebal, isi normal & diurutkan abjad.
 */
import { escapeXml, withProps, RPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { paragraphLine, type VerifyRegion } from "./verify";

const KW_RE = /^\s*((?:kata[\s\u00a0]*kunci|keywords?|key[\s\u00a0]+words)\s*[:：])\s*/i;

export function formatKeywords(
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
): { count: number; regions: VerifyRegion[] } {
  const out = { count: 0, regions: [] as VerifyRegion[] };
  nodes.forEach((n, i) => {
    if (n.kind !== "paragraph" || outline.items[i]?.zone !== "front") return;
    const t = n.text.replace(/\s+/g, " ").trim();
    const m = KW_RE.exec(t);
    if (!m) return;
    const end = /\.\s*$/.test(t) ? "." : "";
    const rest = t.slice(m[0].length).replace(/[.;,]\s*$/, "");
    const sep = rest.includes(";") ? ";" : ",";
    const items = rest.split(sep).map((s) => s.trim()).filter(Boolean);
    if (!items.length) return;
    const sorted = [...items].sort((x, y) => x.localeCompare(y, "id", { sensitivity: "base" }));

    const runs = n.xml.slice(Math.max(0, n.xml.indexOf("</w:pPr>")));
    const firstRun = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/.exec(runs)?.[0] ?? "";
    const rPr = (/<w:rPr>[\s\S]*?<\/w:rPr>/.exec(firstRun)?.[0] ?? "")
      .replace(/<w:(?:b|bCs|i|iCs|u)(?:\s[^>]*)?\/>/g, "")
      .replace(/<w:rStyle\b[^>]*\/>/, "");
    const run = (text: string, bold: boolean) => withProps(
      `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`,
      "w:rPr",
      {
        "w:b": bold ? "<w:b/>" : `<w:b w:val="0"/>`,
        "w:bCs": bold ? "<w:bCs/>" : `<w:bCs w:val="0"/>`,
        "w:i": `<w:i w:val="0"/>`,
        "w:iCs": `<w:iCs w:val="0"/>`,
      },
      RPR_ORDER,
    );

    const label = m[1].replace(/\s+/g, " ");
    const body = sorted.join(`${sep} `) + end;
    const open = /^<w:p\b[^>]*>/.exec(n.xml)?.[0] ?? "<w:p>";
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(n.xml)?.[0] ?? "";
    const before = paragraphLine(n.xml);
    n.xml = `${open}${pPr}${run(label + " ", true)}${run(body, false)}</w:p>`;
    n.text = `${label} ${body}`;
    out.regions.push({ beforeLines: [before].filter(Boolean), afterLines: [paragraphLine(n.xml)].filter(Boolean), beforeText: t, afterText: n.text });
    out.count++;
  });
  return out;
}