/**
 * Jalur REBUILD — hanya untuk sumber PDF.
 *
 * PDF tidak menyimpan tabel, gambar, maupun batas section sebagai struktur,
 * jadi untuk sumber ini dokumen memang harus disusun ulang dari teks. Untuk
 * DOCX jangan pernah memakai jalur ini: pakai `docx/transform.ts` yang
 * menyunting dokumen asli di tempat.
 *
 * Aturan formatnya diambil dari lapisan `rules/` yang sama dengan jalur DOCX,
 * supaya hasil kedua jalur tidak berbeda gaya.
 */

import {
  AlignmentType, Document, Footer, HeadingLevel, LineRuleType, NumberFormat, Packer,
  PageNumber, Paragraph, TableOfContents, TextRun,
} from "docx";
import type { FormatConfig } from "../schemas/format-config";
import { createThesisRules, type Align, type ParaFormat } from "../rules/format-rules";
import type { Outline } from "../structure/outline";
import type { LayoutHint, Role, Zone } from "../structure/roles";

const ALIGN: Record<Align, (typeof AlignmentType)[keyof typeof AlignmentType]> = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  both: AlignmentType.JUSTIFIED,
};

const HEADING: Record<string, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
  Heading1: HeadingLevel.HEADING_1,
  Heading2: HeadingLevel.HEADING_2,
  Heading3: HeadingLevel.HEADING_3,
};

export async function buildDocxFromText(
  paragraphs: string[],
  outline: Outline,
  c: FormatConfig,
): Promise<Buffer> {
  const rules = createThesisRules(c);
  const font = c.fontFamily;
  const size = Math.round(c.fontSize * 2);

  const runs = (text: string, bold?: boolean, upper?: boolean) =>
    (upper ? text.toUpperCase() : text)
      .split("\n")
      .map((t, i) => new TextRun({ text: t, font, size, bold, break: i ? 1 : 0 }));

  const hintOf = (text: string, i: number): LayoutHint =>
    outline.items[i]?.hint ?? {
      align: null, tabCount: (text.match(/\t/g) || []).length, leadingTabsOnly: false,
      listMarker: false, allBold: false, length: text.length, frontKind: null,
    };

  const toPara = (text: string, role: Role, zone: Zone, hint: LayoutHint, withBreak: boolean): Paragraph => {
    const f: ParaFormat = rules.paragraph(role, zone, hint);
    const r = rules.run(role, zone, hint);
    return new Paragraph({
      heading: f.styleId ? HEADING[f.styleId] : undefined,
      alignment: f.align ? ALIGN[f.align] : undefined,
      pageBreakBefore: withBreak,
      keepNext: f.keepNext,
      keepLines: f.keepLines,
      spacing: {
        before: f.before,
        after: f.after,
        ...(f.line ? { line: f.line, lineRule: LineRuleType.AUTO } : {}),
      },
      indent: f.hanging !== undefined
        ? { left: f.left, hanging: f.hanging }
        : f.firstLine
          ? { firstLine: f.firstLine }
          : undefined,
      children: runs(text, r.bold, r.caps),
    });
  };

  const front: (Paragraph | TableOfContents)[] = [];
  if (c.tocMode === "rebuild") {
    front.push(new Paragraph({
      heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER,
      pageBreakBefore: true, spacing: { after: 240 }, children: runs("DAFTAR ISI", true),
    }));
    front.push(new TableOfContents("Daftar Isi", { hyperlink: true, headingStyleRange: "1-3" }));
  }

  const listPage = (title: string, items: string[]) => [
    new Paragraph({
      heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER,
      pageBreakBefore: true, spacing: { after: 240 }, children: runs(title, true),
    }),
    ...items.map((t) => new Paragraph({
      spacing: { line: 240, lineRule: LineRuleType.AUTO, after: 100 }, children: runs(t),
    })),
  ];

  const textsOf = (role: Role) =>
    paragraphs.filter((_, i) => outline.items[i]?.role === role);

  if (c.listFigures) {
    const f = textsOf("figure_caption");
    if (f.length) front.push(...listPage("DAFTAR GAMBAR", f));
  }
  if (c.listTables) {
    const t = textsOf("table_caption");
    if (t.length) front.push(...listPage("DAFTAR TABEL", t));
  }

  const body = paragraphs.map((text, i) => {
    const it = outline.items[i] ?? { role: "body" as Role, zone: "main" as Zone };
    const hint = hintOf(text, i);
    const f = rules.paragraph(it.role, it.zone, hint);
    return toPara(text, it.role, it.zone, hint, Boolean(f.pageBreakBefore) && !it.breakAlreadyPresent);
  });

  const split = Math.min(outline.mainStart, body.length);

  const hs = (id: string, name: string, level: number) => ({
    id, name, basedOn: "Normal", next: "Normal", quickFormat: true,
    paragraph: { outlineLevel: level },
    run: { font, size, bold: true, color: "000000" },
  });

  const align = c.pageNumbers === "bottom-right" ? AlignmentType.RIGHT : AlignmentType.CENTER;
  const pageMain = rules.page("main");
  const margin = {
    top: pageMain.margin.top, bottom: pageMain.margin.bottom,
    left: pageMain.margin.left, right: pageMain.margin.right,
  };
  const paper = { width: pageMain.widthTwips, height: pageMain.heightTwips };

  const footer = c.pageNumbers === "none" || c.pageNumbers === "keep" ? undefined : {
    default: new Footer({
      children: [new Paragraph({
        alignment: align,
        children: [new TextRun({ children: [PageNumber.CURRENT], font, size })],
      })],
    }),
  };

  const sectionProps = (zone: "front" | "main") => ({
    page: {
      size: paper,
      margin,
      pageNumbers: {
        start: 1,
        formatType: rules.numbering(zone) === "lowerRoman"
          ? NumberFormat.LOWER_ROMAN
          : NumberFormat.DECIMAL,
      },
    },
  });

  const sections = split > 0
    ? [
        { properties: sectionProps("front"), footers: footer, children: [...body.slice(0, split), ...front] },
        { properties: sectionProps("main"), footers: footer, children: body.slice(split) },
      ]
    : [{ properties: sectionProps("main"), footers: footer, children: [...front, ...body] }];

  const doc = new Document({
    features: { updateFields: true },
    styles: {
      default: { document: { run: { font, size } } },
      paragraphStyles: [hs("Heading1", "heading 1", 0), hs("Heading2", "heading 2", 1), hs("Heading3", "heading 3", 2)],
    },
    sections,
  });

  return Packer.toBuffer(doc);
}