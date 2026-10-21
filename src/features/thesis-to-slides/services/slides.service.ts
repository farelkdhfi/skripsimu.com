import { askJson } from "@/features/ai/services/openrouter";
import { chunkSystem, deckSystem } from "@/features/ai/prompts";
import { extractParagraphs } from "@/features/file-upload/services/extract-text";
import { chunkSchema, deckSchema, resolveSlideCount, type Deck, type SlidesConfig } from "../schemas/slides";
import { chunkText } from "../utils/concurrency";
import { mapLimit } from "@/lib/concurrency";


const MAX_CHARS = 240_000, DIRECT_LIMIT = 22_000, CHUNK = 12_000;

export async function generateDeck(file: File, c: SlidesConfig): Promise<{ deck: Deck; warnings: string[] }> {
  const { paragraphs } = await extractParagraphs(file);
  const full = paragraphs.join("\n");
  if (full.length < 500) throw Object.assign(new Error("Teks dokumen terlalu sedikit atau tidak terbaca."), { status: 422 });

  const warnings: string[] = [];
  let used = paragraphs;
  if (full.length > MAX_CHARS) {
    warnings.push("Dokumen sangat panjang, hanya bagian awal (±240 ribu karakter) yang dianalisis.");
    used = full.slice(0, MAX_CHARS).split("\n");
  }

  let material: string;
  if (used.join("\n").length <= DIRECT_LIMIT) material = used.join("\n");
  else {
    const chunks = chunkText(used, CHUNK);
    const notes = await mapLimit(chunks, 3, (t) =>
      askJson({ system: chunkSystem, user: `<document>\n${t}\n</document>`, schema: chunkSchema, maxTokens: 1500 }));
    material = notes.map((n) => JSON.stringify(n)).join("\n").slice(0, 40_000);
  }

  const total = resolveSlideCount(c);
  const deck = await askJson({
    system: deckSystem({ total, language: c.language, detail: c.detail, notes: c.speakerNotes }),
    user: `<document>\n${material}\n</document>`,
    schema: deckSchema, maxTokens: 6000,
  });

  if (deck.slides.length > total - 1) deck.slides = deck.slides.slice(0, total - 1);
  if (deck.slides.length < total - 1) warnings.push(`AI hanya menghasilkan ${deck.slides.length + 1} dari ${total} slide yang diminta.`);
  for (const s of deck.slides) {
    const ch = s.visual.chart;
    if (s.visual.kind !== "chart" || !ch || ch.labels.length !== ch.values.length || ch.values.length < 2) {
      if (s.visual.kind === "chart") s.visual = { kind: "none", suggestion: s.visual.suggestion };
      else s.visual.chart = undefined;
    }
  }
  return { deck, warnings };
}