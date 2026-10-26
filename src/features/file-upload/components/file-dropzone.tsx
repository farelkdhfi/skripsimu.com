"use client";
import { useRef, useState } from "react";
import type { KeyboardEvent } from "react";
import { FileText, UploadCloud, X } from "lucide-react";
import { validateFile, MAX_FILE_MB } from "../schemas/file";

const cx = (...c: (string | false | null | undefined)[]) => c.filter(Boolean).join(" ");

const formatSize = (bytes: number) =>
  bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / 1024 / 1024).toFixed(2)} MB`;

export function FileDropzone({ file, onChange }: { file: File | null; onChange: (f: File | null) => void }) {
  const ref = useRef<HTMLInputElement>(null);
  const [drag, setDrag] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const open = () => ref.current?.click();

  const pick = (f?: File) => {
    if (!f) return;
    const err = validateFile(f);
    setError(err);
    if (!err) onChange(f);
  };

  const onKey = (e: KeyboardEvent<HTMLDivElement>) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      open();
    }
  };

  const ext = file?.name.split(".").pop()?.toUpperCase();

  return (
    <div className="space-y-2">
      <div
        onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => { e.preventDefault(); setDrag(false); pick(e.dataTransfer.files[0]); }}
        className={cx("rounded-[1.5rem] transition", drag && "ring-2 ring-neutral-900")}
      >
        <input
          ref={ref}
          type="file"
          accept=".docx,.pdf"
          hidden
          onChange={(e) => {
            pick(e.target.files?.[0]);
            e.target.value = "";
          }}
        />

        {file ? (
          /* Sudah ada file */
          <div className="flex items-center gap-3 rounded-[1.5rem] bg-white p-4">
            <span className="grid size-11 shrink-0 place-items-center rounded-full bg-neutral-900 text-orange-300">
              <FileText className="size-5" />
            </span>
            <div className="min-w-0 flex-1">
              <p className="truncate text-sm font-medium text-neutral-900">{file.name}</p>
              <div className="mt-1 flex items-center gap-1.5">
                {ext && <span className="rounded-full bg-orange-200 px-2 py-0.5 text-[11px] text-neutral-900">{ext}</span>}
                <span className="text-xs text-neutral-500">{formatSize(file.size)}</span>
              </div>
            </div>
            <button
              type="button"
              onClick={open}
              className="shrink-0 rounded-full bg-neutral-100 px-3.5 py-2 text-xs text-neutral-700 transition hover:bg-neutral-200 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-neutral-900"
            >
              Ganti
            </button>
            <button
              type="button"
              onClick={() => { onChange(null); setError(null); }}
              aria-label="Hapus file"
              className="grid size-8 shrink-0 place-items-center rounded-full bg-neutral-100 text-neutral-500 transition hover:bg-neutral-900 hover:text-white focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-neutral-900"
            >
              <X className="size-4" />
            </button>
          </div>
        ) : (
          /* Belum ada file */
          <div
            role="button"
            tabIndex={0}
            aria-label="Pilih file skripsi"
            onClick={open}
            onKeyDown={onKey}
            className={cx(
              "cursor-pointer rounded-[1.5rem] border-2 border-dashed p-8 text-center transition focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-neutral-900",
              drag ? "border-neutral-900 bg-white" : "border-neutral-900/25 bg-white/70 hover:bg-white",
            )}
          >
            <span className="mx-auto grid size-14 place-items-center rounded-full bg-neutral-900 text-orange-300">
              <UploadCloud className="size-6" />
            </span>
            <p className="mt-4 text-sm font-medium text-neutral-900">Tarik file ke sini</p>
            <p className="mt-1 text-xs text-neutral-500">atau</p>
            <span className="mt-2 inline-block rounded-full bg-neutral-900 px-5 py-2 text-xs text-white">Pilih dari perangkat</span>
            <p className="mt-4 text-xs leading-relaxed text-neutral-500">
              DOCX atau PDF, maks {MAX_FILE_MB} MB. Dokumen tidak disimpan permanen.
            </p>
          </div>
        )}
      </div>

      {error && <p className="rounded-2xl bg-red-50 px-4 py-2.5 text-xs leading-relaxed text-red-700">{error}</p>}
    </div>
  );
}