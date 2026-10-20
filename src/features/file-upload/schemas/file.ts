export const MAX_FILE_MB = 15;
export function validateFile(f: File): string | null {
  const n = f.name.toLowerCase();
  if (!n.endsWith(".docx") && !n.endsWith(".pdf")) return "Format harus DOCX atau PDF.";
  if (f.size > MAX_FILE_MB * 1024 * 1024) return `Ukuran file maksimal ${MAX_FILE_MB} MB.`;
  if (f.size === 0) return "File kosong.";
  return null;
}