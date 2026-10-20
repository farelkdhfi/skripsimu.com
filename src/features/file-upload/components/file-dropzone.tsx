"use client";
import { useRef, useState } from "react";
import { FileText, UploadCloud, X } from "lucide-react";
import { validateFile, MAX_FILE_MB } from "../schemas/file";

export function FileDropzone({ file, onChange }: { file: File | null; onChange: (f: File | null) => void }) {
  const ref = useRef<HTMLInputElement>(null);
  const [drag, setDrag] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const pick = (f?: File) => {
    if (!f) return;
    const err = validateFile(f);
    setError(err);
    if (!err) onChange(f);
  };

  return (
    <div>
      <div
        onClick={() => ref.current?.click()}
        onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => { e.preventDefault(); setDrag(false); pick(e.dataTransfer.files[0]); }}
        className={`cursor-pointer rounded-[1.5rem] border-2 border-dashed bg-white/70 p-8 text-center transition ${
          drag ? "border-purple-500 bg-white" : "border-purple-300 hover:bg-white"
        }`}
      >
        <input ref={ref} type="file" accept=".docx,.pdf" hidden onChange={(e) => pick(e.target.files?.[0])} />
        {file ? (
          <div className="flex flex-wrap items-center justify-center gap-3">
            <span className="grid size-10 place-items-center rounded-full bg-green-200 text-slate-900">
              <FileText className="size-5" />
            </span>
            <span className="text-sm font-medium text-slate-900">{file.name}</span>
            <span className="rounded-full bg-slate-100 px-2.5 py-1 text-xs font-medium text-slate-600">{(file.size / 1024 / 1024).toFixed(2)} MB</span>
            <button
              type="button"
              onClick={(e) => { e.stopPropagation(); onChange(null); }}
              aria-label="Hapus file"
              className="grid size-7 place-items-center rounded-full bg-slate-100 transition hover:bg-slate-200"
            >
              <X className="size-4 text-slate-500" />
            </button>
          </div>
        ) : (
          <>
            <span className="mx-auto grid size-14 place-items-center rounded-full bg-purple-200 text-purple-700">
              <UploadCloud className="size-6" />
            </span>
            <p className="mt-3 text-sm font-medium text-slate-900">Tarik file ke sini atau klik untuk memilih</p>
            <p className="mt-1 text-xs text-slate-500">DOCX atau PDF, maks {MAX_FILE_MB} MB. Dokumen tidak disimpan permanen.</p>
          </>
        )}
      </div>
      {error && <p className="mt-2 pl-1 text-sm text-red-600">{error}</p>}
    </div>
  );
}