import { z } from "zod";
import { extractJson } from "../utils/json";

export class AiError extends Error {
  constructor(m: string, public status = 502, public retryable = false) { super(m); }
}

let useJsonMode = true;

export async function askJson<T extends z.ZodTypeAny>(a: {
  system: string; user: string; schema: T; maxTokens?: number;
}): Promise<z.infer<T>> {
  const key = process.env.OPENROUTER_API_KEY;
  if (!key) throw new AiError("OPENROUTER_API_KEY belum diatur di .env.local", 500);

  let lastReason = "";
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
        method: "POST",
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
        // model tidak mendukung response_format -> matikan lalu coba lagi
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
      if (!(e instanceof AiError) || !e.retryable) throw e;
      await new Promise((r) => setTimeout(r, 800 * (attempt + 1)));
    }
  }
  throw new AiError(`AI gagal memberi output yang valid setelah 3 percobaan (${lastReason}).`);
}