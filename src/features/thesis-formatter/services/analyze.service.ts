import { z } from "zod";
import { askJson } from "@/features/ai/services/openrouter";
import { structureSystem, structureUser } from "@/features/ai/prompts";
import { preclassify } from "../utils/preclassify";
import type { Role } from "../types";

type AiRole = "cover" | "chapter" | "section" | "subsection" | "skip" | "body";

function normalizeRole(v: string): AiRole {
  const s = v.toLowerCase().trim();
  if (/cover|sampul|title/.test(s)) return "cover";
  if (/chapter|bab/.test(s)) return "chapter";
  if (/subsection|sub_section|sub-section|level.?3|h3|heading.?3/.test(s)) return "subsection";
  if (/section|subbab|level.?2|h2|heading.?2|heading/.test(s)) return "section";
  if (/skip|toc|list/.test(s)) return "skip";
  return "body";
}

const schema = z.preprocess(
  (v) => (Array.isArray(v) ? { items: v } : v),
  z.object({
    items: z.array(
      z.object({ i: z.coerce.number(), role: z.string().transform(normalizeRole) }).catch({ i: -1, role: "body" }),
    ).catch([]),
    recommendations: z.array(z.string()).catch([]),
  }),
);

export async function classifyParagraphs(paragraphs: string[]) {
  const roles: (Role | null)[] = paragraphs.map(preclassify);
  const cands = paragraphs
    .map((t, i) => ({ i, t }))
    .filter((x) => roles[x.i] === null && x.t.length <= 120 && !/[.!?]$/.test(x.t));

  const recs = new Set<string>();
  let failed = 0;
  for (let s = 0; s < cands.length; s += 60) {
    const batch = cands.slice(s, s + 60);
    try {
      const out = await askJson({ system: structureSystem, user: structureUser(batch), schema, maxTokens: 3000 });
      const valid = new Set(batch.map((b) => b.i));
      for (const it of out.items) if (valid.has(it.i)) roles[it.i] = it.role;
      out.recommendations.forEach((r) => recs.add(r));
    } catch (e) {
      failed++;
      console.error("[classifyParagraphs] batch gagal:", e instanceof Error ? e.message : e);
    }
  }
  if (failed > 0)
    recs.add(`AI gagal menganalisis ${failed} bagian dokumen, sehingga sebagian judul dikenali memakai aturan penomoran saja. Periksa hasil heading.`);
  return { roles, recommendations: [...recs].slice(0, 8) };
}