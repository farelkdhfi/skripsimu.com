import {
  AlignmentType, Document, Footer, HeadingLevel, LineRuleType, Packer, PageNumber,
  Paragraph, TableOfContents, TextRun,
} from "docx";
import type { Block } from "../types";
import type { FormatConfig } from "../schemas/format-config";

const cm = (v: number) => Math.round(v * 567);

export async function buildDocx(blocks: Block[], c: FormatConfig): Promise<Buffer> {
  const font = c.fontFamily, size = c.fontSize * 2;
  const line = { line: Math.round(c.lineSpacing * 240), lineRule: LineRuleType.AUTO };
  const single = { line: 240, lineRule: LineRuleType.AUTO };

  const runs = (text: string, o: { bold?: boolean } = {}) =>
    text.split("\n").map((t, i) => new TextRun({ text: t, font, size, bold: o.bold, break: i ? 1 : 0 }));

  const toPara = (b: Block): Paragraph => {
    switch (b.role) {
      case "cover":
        return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { ...single, after: 120 }, children: runs(b.text, { bold: true }) });
      case "chapter":
        return new Paragraph({
          heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, pageBreakBefore: true,
          spacing: { ...single, after: 240 },
          children: runs(c.chapterUppercase ? b.text.toUpperCase() : b.text, { bold: true }),
        });
      case "section":
        return new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { ...single, before: 240, after: 120 }, children: runs(b.text, { bold: true }) });
      case "subsection":
        return new Paragraph({ heading: HeadingLevel.HEADING_3, keepNext: true, spacing: { ...single, before: 200, after: 120 }, children: runs(b.text, { bold: true }) });
      case "figure_caption":
        return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { ...single, before: 60, after: 240 }, children: runs(b.text) });
      case "table_caption":
        return new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { ...single, before: 240, after: 60 }, children: runs(b.text) });
      case "reference":
        return new Paragraph({ alignment: AlignmentType.LEFT, indent: { left: cm(1), hanging: cm(1) }, spacing: { ...single, after: 120 }, children: runs(b.text) });
      default:
        return new Paragraph({ alignment: AlignmentType.JUSTIFIED, indent: { firstLine: cm(c.indent) }, spacing: { ...line, after: 0 }, children: runs(b.text) });
    }
  };

  const listPages = (title: string, items: string[]) => [
    new Paragraph({ alignment: AlignmentType.CENTER, pageBreakBefore: true, spacing: { after: 240 }, children: runs(title, { bold: true }) }),
    ...items.map((t) => new Paragraph({ spacing: { ...single, after: 100 }, children: runs(t) })),
  ];

  const coverEnd = blocks.findIndex((b) => b.role !== "cover");
  const split = coverEnd === -1 ? blocks.length : coverEnd;
  const front: (Paragraph | TableOfContents)[] = [];
  if (c.toc) {
    front.push(new Paragraph({ alignment: AlignmentType.CENTER, pageBreakBefore: true, spacing: { after: 240 }, children: runs("DAFTAR ISI", { bold: true }) }));
    front.push(new TableOfContents("Daftar Isi", { hyperlink: true, headingStyleRange: "1-3" }));
  }
  if (c.listFigures) { const f = blocks.filter((b) => b.role === "figure_caption").map((b) => b.text); if (f.length) front.push(...listPages("DAFTAR GAMBAR", f)); }
  if (c.listTables) { const t = blocks.filter((b) => b.role === "table_caption").map((b) => b.text); if (t.length) front.push(...listPages("DAFTAR TABEL", t)); }

  const hs = (id: string, name: string) => ({
    id, name, basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { font, size, bold: true, color: "000000" },
  });
  const align = c.pageNumbers === "bottom-right" ? AlignmentType.RIGHT : AlignmentType.CENTER;
  const paper = c.paper === "A4" ? { width: 11906, height: 16838 } : { width: 12240, height: 15840 };

  const doc = new Document({
    features: { updateFields: true },
    styles: {
      default: { document: { run: { font, size } } },
      paragraphStyles: [hs("Heading1", "heading 1"), hs("Heading2", "heading 2"), hs("Heading3", "heading 3")],
    },
    sections: [{
      properties: { page: { size: paper, margin: { top: cm(c.marginTop), bottom: cm(c.marginBottom), left: cm(c.marginLeft), right: cm(c.marginRight) } } },
      footers: c.pageNumbers === "none" ? undefined : {
        default: new Footer({ children: [new Paragraph({ alignment: align, children: [new TextRun({ children: [PageNumber.CURRENT], font, size })] })] }),
      },
      children: [...blocks.slice(0, split).map(toPara), ...front, ...blocks.slice(split).map(toPara)],
    }],
  });
  return Packer.toBuffer(doc);
}