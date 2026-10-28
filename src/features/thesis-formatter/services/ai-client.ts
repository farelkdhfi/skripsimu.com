/**
 * Panggilan AI dari browser: hanya TEKS kecil yang dikirim ke server (/api/formatter/ai), bukan
 * berkas naskah. Gagal → hasil kosong (pembacaan aturan tetap jalan), sama seperti di server.
 */
import type { AiPorts } from "./format-core";
import type { Signals } from "../structure/roles";

const URL_AI = "/api/formatter/ai";
/** Teks per paragraf dipotong: aturan & AI struktur hanya butuh awal paragraf. Menjaga kiriman jauh di bawah 4,5 MB. */
const MAX_TEXT = 600;

async function call<T>(kind: string, payload: unknown, fallback: T): Promise<T> {
  try {
    const res = await fetch(URL_AI, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ kind, payload }),
    });
    if (!res.ok) {
      console.warn(`[ai-client] ${kind}: HTTP ${res.status}`);
      return fallback;
    }
    return (await res.json()) as T;
  } catch (e) {
    console.warn(`[ai-client] ${kind} gagal:`, e instanceof Error ? e.message : e);
    return fallback;
  }
}

export const browserAi: AiPorts = {
  classify: (signals: Signals[]) => call("classify", signals.map((s) => ({ ...s, text: s.text.slice(0, MAX_TEXT) })), {
    roles: signals.map(() => null), recommendations: [], notices: ["Analisis AI struktur tidak tersedia; dipakai aturan saja."],
  }),
  cover: (lines: string[]) => call("cover", lines, null),
  front: (listing: string) => call("front", listing, null),
};