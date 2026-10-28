/**
 * Pemformatan DOCX di browser. Dijalankan di Web Worker; bila Worker tidak tersedia, langsung di
 * halaman (lebih lambat terasa, hasil sama).
 */
import type { FormatConfig } from "../schemas/format-config";
import type { FormatResult } from "../types";
import type { FrontFieldInput } from "../front/types";
import type { WorkerRequest, WorkerResponse } from "../workers/format.worker";

export async function formatDocxInBrowser(file: File, config: FormatConfig, frontFields?: FrontFieldInput): Promise<FormatResult> {
  const bytes = await file.arrayBuffer();
  const req: WorkerRequest = { bytes, fileName: file.name, config, frontFields };

  let worker: Worker | null = null;
  try {
    worker = new Worker(new URL("../workers/format.worker.ts", import.meta.url), { type: "module" });
  } catch (e) {
    console.warn("[format-browser] Web Worker tidak tersedia, diproses di halaman:", e);
  }

  if (!worker) {
    const [{ formatDocxBytes }, { browserAi }] = await Promise.all([import("./format-core"), import("./ai-client")]);
    return formatDocxBytes(new Uint8Array(bytes), file.name, config, frontFields, browserAi);
  }

  return new Promise<FormatResult>((resolve, reject) => {
    worker!.onmessage = (ev: MessageEvent<WorkerResponse>) => {
      worker!.terminate();
      if (ev.data.ok) resolve(ev.data.result);
      else reject(Object.assign(new Error(ev.data.error), { status: 422 }));
    };
    worker!.onerror = (ev) => {
      worker!.terminate();
      reject(new Error(ev.message || "Pemformatan di browser gagal."));
    };
    // ArrayBuffer dipindahkan (bukan disalin) ke worker.
    worker!.postMessage(req, [bytes]);
  });
}