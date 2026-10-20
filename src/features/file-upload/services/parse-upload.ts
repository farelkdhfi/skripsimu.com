import { validateFile } from "../schemas/file";

export async function parseUpload(req: Request) {
  const fd = await req.formData();
  const file = fd.get("file");
  if (!(file instanceof File)) throw Object.assign(new Error("File tidak ditemukan."), { status: 400 });
  const err = validateFile(file);
  if (err) throw Object.assign(new Error(err), { status: 400 });
  return { file, fd };
}