export async function mapLimit<T, R>(items: T[], limit: number, fn: (x: T) => Promise<R>): Promise<R[]> {
  const out: R[] = new Array(items.length);
  let next = 0;
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) { const i = next++; out[i] = await fn(items[i]); }
  }));
  return out;
}
export function chunkText(paragraphs: string[], size: number) {
  const chunks: string[] = []; let cur = "";
  for (const p of paragraphs) {
    if (cur.length + p.length > size && cur) { chunks.push(cur); cur = ""; }
    cur += p + "\n";
  }
  if (cur) chunks.push(cur);
  return chunks;
}