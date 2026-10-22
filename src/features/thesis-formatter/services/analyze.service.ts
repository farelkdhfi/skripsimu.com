import { z } from "zod";
import { askJson } from "@/features/ai/services/openrouter";
import { structureSystem, structureUser } from "@/features/ai/prompts";
import { mapLimit } from "@/lib/concurrency";
import { classifyByRule, scanStructure, type Role, type Signals } from "../structure/roles";

const BATCH = 150;
const MAX_BATCHES = 5;
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
  if (i < 30) return t.length <= 120;
  return (
    t.length >= 3 && t.length <= 100 &&
    t.split(/\s+/).length <= 14 &&
    !/[.!?;,:]$/.test(t) &&
    !/[\s.]\d{1,3}$/.test(t) &&
    !/^[a-z]/.test(t)
  );
}

export type ClassifyResult = {
  /** Peran per indeks node; `null` berarti AI tidak memutuskan apa pun. */
  roles: (Role | null)[];
  recommendations: string[];
  notices: string[];
};

/**
 * AI hanya dipanggil untuk node yang TIDAK bisa diputuskan oleh aturan.
 * Penanda struktural yang sudah pasti (BAB, 1.1, DAFTAR PUSTAKA, keterangan
 * gambar) tidak pernah ditimpa oleh model, sehingga urutan dan susunan
 * dokumen tidak bisa berubah gara-gara tebakan AI.
 */
export async function classifyNodes(signals: Signals[]): Promise<ClassifyResult> {
  const { zones, inList } = scanStructure(signals);
  const roles: (Role | null)[] = signals.map((s, i) => {
    const r = classifyByRule(s, zones[i], inList[i]);
    return r === "empty" ? null : r;
  });

  const cands = signals
    .map((s, i) => ({ i, t: s.text.slice(0, 100) }))
    .filter((x) => !signals[x.i].isTable && roles[x.i] === null && x.t.trim() && isCandidate(signals[x.i].text, x.i));

  const batches: { n: number; items: { i: number; t: string }[] }[] = [];
  for (let s = 0; s < cands.length; s += BATCH) batches.push({ n: batches.length, items: cands.slice(s, s + BATCH) });

  const deadline = Date.now() + (Number(process.env.AI_TIME_BUDGET_MS) || 100_000);
  const recs = new Set<string>();
  const unanalyzed: number[] = [];
  let budgetSkips = 0;
  let failures = 0;
  const out: (Role | null)[] = signals.map(() => null);

  await mapLimit(batches, CONCURRENCY, async ({ n, items }) => {
    if (n >= MAX_BATCHES || Date.now() > deadline - 3000) {
      unanalyzed.push(items[0].i);
      budgetSkips++;
      return;
    }
    try {
      const res = await askJson({
        system: structureSystem, user: structureUser(items), schema,
        maxTokens: 2000, timeoutMs: 40_000, maxAttempts: 2, deadlineAt: deadline,
      });
      const valid = new Set(items.map((b) => b.i));
      for (const [i, code] of res.h) {
        const role = toRole(code);
        if (role && valid.has(i)) out[i] = role;
      }
      res.r.slice(0, 3).forEach((r) => recs.add(r));
    } catch (e) {
      unanalyzed.push(items[0].i);
      failures++;
      console.error("[classifyNodes] batch gagal:", e instanceof Error ? e.message : e);
    }
  });

  const notices: string[] = [];
  if (unanalyzed.length) {
    const first = Math.min(...unanalyzed);
    let heading = "awal dokumen";
    for (let k = first - 1; k >= 0; k--) {
      const r = roles[k];
      if (r === "chapter" || r === "chapter_label" || r === "section") {
        heading = `"${signals[k].text.slice(0, 60)}"`;
        break;
      }
    }
    const cause = failures && !budgetSkips ? "layanan AI tidak dapat dihubungi"
      : budgetSkips && !failures ? "batas waktu/token"
      : "batas waktu/token dan kegagalan layanan AI";
    notices.push(
      `AI belum menganalisis ${unanalyzed.length} bagian dokumen karena ${cause}, dimulai dari sekitar ${heading} (paragraf ke-${first + 1}). ` +
      `Judul bernomor tetap dikenali aturan, dan struktur dokumen asli tetap dipertahankan. Font, spasi, margin, tabel, dan gambar tetap diproses seperti biasa.`,
    );
  }

  return { roles: out, recommendations: [...recs].slice(0, 8), notices };
}