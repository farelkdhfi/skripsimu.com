import PptxGenJS from "pptxgenjs/dist/pptxgen.cjs.js";
import { THEMES, type ThemeKey } from "../utils/themes";
import type { Deck } from "../schemas/slides";

export async function buildPptx(deck: Deck, style: ThemeKey): Promise<Buffer> {
  const t = THEMES[style];
  const p = new PptxGenJS();
  p.layout = "LAYOUT_16x9";
  p.title = deck.title;

  const s0 = p.addSlide();
  s0.background = { color: t.titleBg };
  s0.addText(deck.title, { x: 0.7, y: 1.2, w: 8.6, h: 1.8, fontFace: t.font, fontSize: 32, bold: true, color: t.titleText, valign: "bottom" });
  s0.addShape(p.ShapeType.rect, { x: 0.7, y: 3.15, w: 1.2, h: 0.07, fill: { color: t.accent }, line: { color: t.accent } });
  if (deck.subtitle) s0.addText(deck.subtitle, { x: 0.7, y: 3.35, w: 8.6, h: 1, fontFace: t.font, fontSize: 16, color: t.titleText, valign: "top" });

  for (const sl of deck.slides) {
    const s = p.addSlide();
    s.background = { color: t.bg };
    s.addShape(p.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: t.accent }, line: { color: t.accent } });
    s.addText(sl.title, { x: 0.6, y: 0.35, w: 8.8, h: 0.8, fontFace: t.font, fontSize: 26, bold: true, color: t.heading, valign: "middle" });

    const ch = sl.visual.kind === "chart" ? sl.visual.chart : undefined;
    const fs = sl.bullets.length > 5 ? 15 : 18;
    if (sl.bullets.length)
      s.addText(sl.bullets.map((b) => ({ text: b, options: { bullet: true, breakLine: true } })), {
        x: 0.6, y: 1.35, w: ch ? 5.1 : 8.8, h: 3.7, fontFace: t.font, fontSize: fs, color: t.text, valign: "top", paraSpaceAfter: 8,
      });
    if (ch) {
      const type = ch.type === "line" ? p.ChartType.line : ch.type === "pie" ? p.ChartType.pie : p.ChartType.bar;
      s.addChart(type, [{ name: ch.title ?? "Data", labels: ch.labels, values: ch.values }], {
        x: 5.9, y: 1.4, w: 3.6, h: 3.5, showLegend: ch.type === "pie", legendPos: "b",
        chartColors: [t.accent, "A5B4FC", "6366F1", "C7D2FE", "4338CA"], showTitle: !!ch.title, title: ch.title,
        catAxisLabelColor: t.muted, valAxisLabelColor: t.muted, catAxisLabelFontSize: 10, valAxisLabelFontSize: 10,
      });
    }
    const notes = [sl.notes, sl.visual.suggestion && `[Saran visual] ${sl.visual.suggestion}`].filter(Boolean).join("\n\n");
    if (notes) s.addNotes(notes);
    s.slideNumber = { x: 9.2, y: 5.2, fontSize: 10, color: t.muted };
  }
  return (await p.write({ outputType: "nodebuffer" })) as Buffer;
}