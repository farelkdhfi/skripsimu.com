import { z } from "zod";

export const slidesConfigSchema = z.object({
  preset: z.enum(["5", "10", "15", "20", "custom"]),
  customCount: z.number().int().min(3, "Minimal 3").max(40, "Maksimal 40"),
  language: z.enum(["Indonesia", "English"]),
  style: z.enum(["academic-blue", "soft-violet", "mono-dark"]),
  detail: z.enum(["ringkas", "sedang", "detail"]),
  speakerNotes: z.boolean(),
});
export type SlidesConfig = z.infer<typeof slidesConfigSchema>;
export const resolveSlideCount = (c: SlidesConfig) => (c.preset === "custom" ? c.customCount : Number(c.preset));

const chart = z.object({
  type: z.enum(["bar", "line", "pie"]), title: z.string().optional(),
  labels: z.array(z.string()), values: z.array(z.number()),
});
const slide = z.object({
  title: z.string().min(1),
  bullets: z.array(z.string()).default([]),
  notes: z.string().default(""),
  visual: z.object({
    kind: z.enum(["none", "chart", "table", "image"]).default("none"),
    suggestion: z.string().optional(),
    chart: chart.optional(),
  }).default({ kind: "none" }),
});
export const deckSchema = z.object({ title: z.string(), subtitle: z.string().default(""), slides: z.array(slide).min(1) });
export const chunkSchema = z.object({
  section: z.string().default(""), facts: z.array(z.string()).default([]),
  data: z.array(z.object({ label: z.string(), value: z.number(), context: z.string().default("") })).default([]),
});
export type Deck = z.infer<typeof deckSchema>;