import { useMutation } from "@tanstack/react-query";
import { postForm } from "@/lib/api-client";
import { resolveSlideCount, type Deck, type SlidesConfig } from "../schemas/slides";

export function useGenerateSlides() {
  return useMutation({
    mutationFn: ({ file, config }: { file: File; config: SlidesConfig }) => {
      const fd = new FormData();
      fd.append("file", file);
      fd.append("config", JSON.stringify({ ...config, slideCount: resolveSlideCount(config) }));
      return postForm<{ deck: Deck; warnings: string[] }>("/api/slides/generate", fd);
    },
  });
}

export function useExportPptx() {
  return useMutation({
    mutationFn: async ({ deck, style }: { deck: Deck; style: SlidesConfig["style"] }) => {
      const res = await fetch("/api/slides/export", {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ deck, style }),
      });
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error ?? "Gagal mengekspor PPTX");
      return res.blob();
    },
  });
}