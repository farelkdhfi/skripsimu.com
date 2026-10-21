import { z } from "zod";
import { extractJson } from "../utils/json";

export class AiError extends Error {
  constructor(m: string, public status = 502, public retryable = false) { super(m); }
}

let useJsonMode = true;

export async function askJson<T extends z.ZodTypeAny>(a: {
  system: string; user: string; schema: T; maxTokens?: number;
  timeoutMs?: number; maxAttempts?: number; deadlineAt?: number;
}): Promise<z.infer<T>> {
  const key = process.env.OPENROUTER_API_KEY;
  if (!key) throw new AiError("OPENROUTER_API_KEY belum diatur di environment variable", 500);

  const attempts = a.maxAttempts ?? 3;
  let lastReason = "";
  for (let attempt = 0; attempt < attempts; attempt++) {
    const remaining = a.deadlineAt ? a.deadlineAt - Date.now() : Infinity;
    if (remaining < 2000) throw new AiError("Batas waktu proses AI tercapai", 504, false);
    const timeoutMs = Math.min(a.timeoutMs ?? 90_000, remaining);

    try {
      const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
        signal: AbortSignal.timeout(timeoutMs),
        headers: { Authorization: `Bearer ${key}`, "Content-Type": "application/json", "X-Title": "SkripsiAI" },
        body: JSON.stringify({
          model: process.env.OPENROUTER_MODEL || "openai/gpt-4o-mini",
          temperature: 0.2,
          max_tokens: a.maxTokens ?? 4000,
          ...(useJsonMode ? { response_format: { type: "json_object" } } : {}),
          messages: [{ role: "system", content: a.system }, { role: "user", content: a.user }],
        }),
      });

      if (!res.ok) {
        const t = await res.text();
        if (res.status === 400 && /response_format|json_object|structured/i.test(t) && useJsonMode) {
          useJsonMode = false;
          throw new AiError("JSON mode tidak didukung model, mencoba tanpa JSON mode", 502, true);
        }
        throw new AiError(`OpenRouter ${res.status}: ${t.slice(0, 200)}`, 502, res.status === 429 || res.status >= 500);
      }

      const data = await res.json();
      const choice = data?.choices?.[0];
      let content = choice?.message?.content;
      if (Array.isArray(content)) content = content.map((p: { text?: string }) => p?.text ?? "").join("");
      if (typeof content !== "string" || !content.trim()) {
        lastReason = `respons kosong (finish_reason=${choice?.finish_reason ?? "?"})`;
        throw new AiError(lastReason, 502, true);
      }
      if (choice?.finish_reason === "length") {
        lastReason = "output terpotong (max_tokens habis)";
        console.error("[AI]", lastReason, content.slice(-120));
        throw new AiError(lastReason, 502, true);
      }

      let parsed: unknown;
      try {
        parsed = JSON.parse(extractJson(content));
      } catch {
        lastReason = "bukan JSON valid";
        console.error("[AI] JSON parse gagal. Cuplikan:", content.slice(0, 300));
        throw new AiError(lastReason, 502, true);
      }
      const r = a.schema.safeParse(parsed);
      if (!r.success) {
        lastReason = "schema tidak cocok: " + r.error.issues.slice(0, 3).map((i) => `${i.path.join(".")} ${i.message}`).join("; ");
        console.error("[AI]", lastReason, "\nCuplikan:", content.slice(0, 300));
        throw new AiError(lastReason, 502, true);
      }
      return r.data;
    } catch (e) {
      const isTimeout = e instanceof Error && (e.name === "TimeoutError" || e.name === "AbortError");
      const err = isTimeout ? new AiError(`Request AI melebihi batas waktu ${Math.round(timeoutMs / 1000)} detik`, 504, false) : e;
      if (!(err instanceof AiError) || !err.retryable) throw err;
      await new Promise((r) => setTimeout(r, 800 * (attempt + 1)));
    }
  }
  throw new AiError(`AI gagal memberi output yang valid setelah ${attempts} percobaan (${lastReason}).`);
}