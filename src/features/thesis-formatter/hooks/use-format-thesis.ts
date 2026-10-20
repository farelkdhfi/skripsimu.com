import { useMutation } from "@tanstack/react-query";
import { postForm } from "@/lib/api-client";
import type { FormatConfig } from "../schemas/format-config";
import type { FormatResult } from "../types";

export function useFormatThesis() {
  return useMutation({
    mutationFn: ({ file, config }: { file: File; config: FormatConfig }) => {
      const fd = new FormData();
      fd.append("file", file);
      fd.append("config", JSON.stringify(config));
      return postForm<FormatResult>("/api/formatter", fd);
    },
  });
}