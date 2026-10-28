import { useMutation } from "@tanstack/react-query";
import { postForm } from "@/lib/api-client";
import type { FormatConfig } from "../schemas/format-config";
import type { FormatResult } from "../types";
import type { FrontFieldInput } from "../front/types";

/**
 * DOCX diformat DI BROWSER (Web Worker): naskah tidak diunggah, jadi tidak terkena batas ukuran
 * request Vercel (4,5 MB) — skripsi 10–100 MB tetap bisa. Server hanya dipanggil untuk AI (teks kecil).
 * PDF masih lewat server (jalur susun ulang).
 */
export function useFormatThesis() {
  return useMutation({
    mutationFn: async ({ file, config, frontFields }: { file: File; config: FormatConfig; frontFields?: FrontFieldInput }) => {
      if (/\.docx$/i.test(file.name)) {
        const { formatDocxInBrowser } = await import("../services/format-browser");
        return formatDocxInBrowser(file, config, frontFields);
      }
      const fd = new FormData();
      fd.append("file", file);
      fd.append("config", JSON.stringify(config));
      if (frontFields) fd.append("frontFields", JSON.stringify(frontFields));
      return postForm<FormatResult>("/api/formatter", fd);
    },
  });
}