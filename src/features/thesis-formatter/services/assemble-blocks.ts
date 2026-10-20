import type { Block, Role } from "../types";
import type { FormatConfig } from "../schemas/format-config";

const REF = /daftar pustaka|referensi|references|bibliography/i;

export function assembleBlocks(paragraphs: string[], roles: (Role | null)[], c: FormatConfig) {
  const blocks: Block[] = [];
  let removed = 0, inRef = false, skipping = false, coverDone = false;

  for (let i = 0; i < paragraphs.length; i++) {
    let text = paragraphs[i];
    let role: Role = roles[i] ?? "body";

    if (role === "chapter") {
      const next = paragraphs[i + 1];
      if (/^bab\s+[ivxlc\d]+\.?$/i.test(text) && next && next.length <= 90 && !/\.$/.test(next) &&
          [null, "chapter", "body"].includes(roles[i + 1] as Role | null)) {
        text += `\n${paragraphs[++i]}`;
      }
      if (c.toc && /^daftar\s+(isi|gambar|tabel)/i.test(text)) { skipping = true; removed++; continue; }
      if (skipping && !/(\t|\s)\d+$|\.{3,}/.test(text)) skipping = false;
    }
    if (skipping && text.length > 150) skipping = false;
    if (skipping || role === "skip") { removed++; continue; }

    if (role !== "cover") coverDone = true;
    else if (coverDone) role = "body";
    if (role === "chapter") inRef = REF.test(text);
    if (role === "body" && inRef) role = "reference";
    blocks.push({ role, text });
  }

  const n = (r: Role) => blocks.filter((b) => b.role === r).length;
  const words = (arr: string[]) => arr.join(" ").split(/\s+/).filter(Boolean).length;
  const stats = {
    chapters: n("chapter"), sections: n("section"), subsections: n("subsection"),
    figures: n("figure_caption"), tables: n("table_caption"), references: n("reference"), removed,
    wordsBefore: words(paragraphs), wordsAfter: words(blocks.map((b) => b.text)),
  };
  return { blocks, stats };
}