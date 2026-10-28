import { askJson } from "@/features/ai/services/openrouter";
import { FRONT_AI_SYSTEM, frontAiSchema, type FrontPointers } from "../front/ai";

/**
 * AI penunjuk data bagian depan. AI hanya menjawab NOMOR baris/gambar (dan potongan teks yang
 * harus ada di baris itu); isi yang dipakai template selalu disalin dari naskah dan diperiksa
 * aturan di front/ai.ts. Gagal/timeout → null, pembacaan aturan yang dipakai.
 */
export async function extractFrontAI(listing: string): Promise<FrontPointers | null> {
  if (!listing.trim()) return null;
  try {
    return await askJson({
      system: FRONT_AI_SYSTEM,
      user: listing,
      schema: frontAiSchema,
      maxTokens: 1200,
      timeoutMs: 30_000,
      maxAttempts: 2,
    });
  } catch (e) {
    console.error("[extractFrontAI]", e instanceof Error ? e.message : e);
    return null;
  }
}