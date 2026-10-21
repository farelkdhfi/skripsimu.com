
export function chunkText(paragraphs: string[], size: number) {
  const chunks: string[] = []; let cur = "";
  for (const p of paragraphs) {
    if (cur.length + p.length > size && cur) { chunks.push(cur); cur = ""; }
    cur += p + "\n";
  }
  if (cur) chunks.push(cur);
  return chunks;
}