import { z } from "zod";
import { askJson } from "@/features/ai/services/openrouter";
import { structureSystem, structureUser } from "@/features/ai/prompts";
import { mapLimit } from "@/lib/concurrency";
import { preclassify } from "../utils/preclassify";
import type { Role } from "../types";

const BATCH = 150;        // baris per request AI
const MAX_BATCHES = 5;    // maks ±750 baris kandidat per dokumen
const CONCURRENCY = 3;

const CODES: Record<string, Role> = { c: "chapter", s: "section", u: "subsection", v: "cover" };

function toRole(v: string): Role | null {
  const s = v.toLowerCase().trim();
  if (CODES[s]) return CODES[s];
  if (/cover|sampul/.test(s)) return "cover";
  if (/subsection|sub_section|sub-section|level.?3|h3/.test(s)) return "subsection";
  if (/chapter|bab/.test(s)) return "chapter";
  if (/section|subbab|level.?2|h2|heading/.test(s)) return "section";
  return null;
}

// Toleran: terima {h:[[i,"c"]]}, array langsung, atau objek {i, role}
const schema = z.preprocess(
  (v) => {
    const o = (v && typeof v === "object" && !Array.isArray(v) ? v : {}) as Record<string, unknown>;
    const raw = (Array.isArray(v) ? v : (o.h ?? o.items ?? [])) as unknown[];
    const h = (Array.isArray(raw) ? raw : [])
      .map((x) => {
        if (Array.isArray(x)) return x;
        if (x && typeof x === "object") {
          const r = x as Record<string, unknown>;
          return [r.i, r.role ?? r.r];
        }
        return null;
      })
      .filter(Boolean);
    return { h, r: o.r ?? o.recommendations ?? [] };
  },
  z.object({
    h: z.array(z.tuple([z.coerce.number(), z.string()]).catch([-1, ""])).catch([]),
    r: z.array(z.string()).catch([]),
  }),
);

function isCandidate(t: string, i: number) {
  if (i < 30) return t.length <= 120; // area sampul
  return (
    t.length >= 3 && t.length <= 100 &&
    t.split(/\s+/).length <= 14 &&
    !/[.!?;,:]$/.test(t) &&
    !/[\s.]\d{1,3}$/.test(t) && // sisa daftar isi (berakhir nomor halaman)
    !/^[a-z]/.test(t)
  );
}

export async function classifyParagraphs(paragraphs: string[]) {
  const roles: (Role | null)[] = paragraphs.map(preclassify);
  const cands = paragraphs
    .map((t, i) => ({ i, t: t.slice(0, 100) }))
    .filter((x) => roles[x.i] === null && isCandidate(paragraphs[x.i], x.i));

  const batches = [];
  for (let s = 0; s < cands.length; s += BATCH) batches.push({ n: batches.length, items: cands.slice(s, s + BATCH) });

  const deadline = Date.now() + (Number(process.env.AI_TIME_BUDGET_MS) || 100_000);
  const recs = new Set<string>();
  const unanalyzed: number[] = []; // indeks paragraf pertama dari tiap batch yang tidak terproses

  await mapLimit(batches, CONCURRENCY, async ({ n, items }) => {
    if (n >= MAX_BATCHES || Date.now() > deadline - 3000) { unanalyzed.push(items[0].i); return; }
    try {
      const out = await askJson({
        system: structureSystem, user: structureUser(items), schema,
        maxTokens: 2000, timeoutMs: 40_000, maxAttempts: 2, deadlineAt: deadline,
      });
      const valid = new Set(items.map((b) => b.i));
      for (const [i, code] of out.h) {
        const role = toRole(code);
        if (role && valid.has(i)) roles[i] = role;
      }
      out.r.slice(0, 3).forEach((r) => recs.add(r));
    } catch (e) {
      unanalyzed.push(items[0].i);
      console.error("[classifyParagraphs] batch gagal:", e instanceof Error ? e.message : e);
    }
  });

  const notices: string[] = [];
  if (unanalyzed.length) {
    const first = Math.min(...unanalyzed);
    let heading = "awal dokumen";
    for (let k = first - 1; k >= 0; k--) {
      if (roles[k] === "chapter" || roles[k] === "section") { heading = `"${paragraphs[k].slice(0, 60)}"`; break; }
    }
    notices.push(
      `AI belum menganalisis ${unanalyzed.length} bagian dokumen karena batas waktu/token, dimulai dari sekitar ${heading} (paragraf ke-${first + 1}). ` +
      `Judul bernomor tetap dikenali aturan; sisanya ditampilkan apa adanya tanpa penyesuaian heading. Font, spasi, dan margin tetap diterapkan.`,
    );
  }
  return { roles, recommendations: [...recs].slice(0, 8), notices };
}