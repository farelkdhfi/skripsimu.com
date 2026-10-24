/**
 * Baris rumus:
 *  - bernomor ("rumus … (2.1)"): rumus di tengah lewat tab tengah, nomor rata kanan
 *    lewat tab kanan, tanpa indentasi
 *  - tanpa nomor (hanya rumus): rata tengah tanpa indentasi
 * Rumus di tengah kalimat (inline) tidak disentuh.
 */
import { withProps, PPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";

const EQ_NUM_RE = /^\(\s*[A-Za-z]?\d+(?:[.\-]\d+)*[a-z]?\s*\)$/;
const wText = (xml: string) =>
  [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]).join("").replace(/\s+/g, " ").trim();

export function layoutEquations(nodes: (ParaNode | TableNode)[], width: number): { numbered: number; centered: number } {
  const out = { numbered: 0, centered: 0 };
  for (const n of nodes) {
    if (n.kind !== "paragraph" || n.hasNumbering || n.hasDrawing) continue;
    if (!/<m:oMath\b/.test(n.xml)) continue;
    const visible = wText(n.xml);

    if (!visible) {
      n.xml = withProps(n.xml, "w:pPr", {
        "w:ind": `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`,
        "w:jc": `<w:jc w:val="center"/>`,
      }, PPR_ORDER);
      out.centered++;
      continue;
    }
    if (!EQ_NUM_RE.test(visible) || /<m:oMathPara\b/.test(n.xml)) continue;

    let xml = n.xml;
    const pEnd = Math.max(0, xml.indexOf("</w:pPr>"));
    const first = xml.indexOf("<m:oMath", pEnd);
    // Tab sebelum rumus (rumus ke posisi tab tengah).
    if (!/<w:tab\/>/.test(xml.slice(pEnd, first))) {
      xml = xml.slice(0, first) + `<w:r><w:tab/></w:r>` + xml.slice(first);
    }
    // Tab sebelum nomor (nomor ke posisi tab kanan).
    const lastEnd = xml.lastIndexOf("</m:oMath>") + "</m:oMath>".length;
    if (!/<w:tab\/>/.test(xml.slice(lastEnd))) {
      xml = xml.slice(0, lastEnd) + `<w:r><w:tab/></w:r>` + xml.slice(lastEnd);
    }
    n.xml = withProps(xml, "w:pPr", {
      "w:keepLines": "<w:keepLines/>",
      "w:tabs": `<w:tabs><w:tab w:val="center" w:pos="${Math.round(width / 2)}"/><w:tab w:val="right" w:pos="${width}"/></w:tabs>`,
      "w:ind": `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`,
      "w:jc": `<w:jc w:val="left"/>`,
    }, PPR_ORDER);
    out.numbered++;
  }
  return out;
}