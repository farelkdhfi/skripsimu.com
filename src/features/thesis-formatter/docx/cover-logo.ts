/**
 * Gambar di halaman sampul (logo) selalu tepat di tengah secara horizontal,
 * walau penulis mendorongnya dengan spasi/tab atau memakai gambar floating.
 */
import { withProps, PPR_ORDER } from "./xml";
import type { ParaNode, TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import { paragraphLine, type VerifyRegion } from "./verify";

const OBJECT_RE = /<w:drawing\b|<w:pict\b|<mc:AlternateContent\b|<w:object\b/;

export function centerCoverImages(
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
): { centered: number; regions: VerifyRegion[] } {
  const out = { centered: 0, regions: [] as VerifyRegion[] };
  nodes.forEach((p, i) => {
    if (p.kind !== "paragraph" || !p.hasDrawing || outline.items[i]?.zone !== "cover") return;
    if (p.text.trim()) return;                    // logo bersama teks: jangan diutak-atik
    if (p.xml.includes("<w:txbxContent")) return; // kotak teks: struktur run bersarang

    const before = paragraphLine(p.xml);
    let xml = p.xml.replace(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g, (run) => {
      if (OBJECT_RE.test(run)) {
        // Run gambar: buang tab/spasi yang ikut menempel di run yang sama.
        return run.replace(/<w:tab\/>/g, "").replace(/<w:t(?:\s[^>]*)?>\s*<\/w:t>/g, "");
      }
      const inner = run.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, "");
      const hasText = /<w:t(?:\s[^>]*)?>[^<]*\S[^<]*<\/w:t>/.test(inner);
      const keep = /<w:(br|cr|fldChar|instrText|sym)\b/.test(inner);
      return hasText || keep ? run : "";          // run berisi spasi/tab saja → dibuang
    });

    xml = withProps(xml, "w:pPr", {
      "w:tabs": null,
      "w:ind": `<w:ind w:left="0" w:right="0" w:firstLine="0"/>`,
      "w:jc": `<w:jc w:val="center"/>`,
    }, PPR_ORDER);

    // Gambar floating: posisi horizontal dikunci ke tengah margin.
    xml = xml.replace(
      /<wp:positionH\b[^>]*>[\s\S]*?<\/wp:positionH>/g,
      `<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>`,
    );

    if (xml === p.xml) return;
    p.xml = xml;
    out.centered++;
    const after = paragraphLine(p.xml);
    if (before !== after) {
      out.regions.push({ beforeLines: [before].filter(Boolean), afterLines: [after].filter(Boolean), beforeText: "", afterText: "" });
    }
  });
  return out;
}