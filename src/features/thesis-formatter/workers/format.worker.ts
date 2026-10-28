/// <reference lib="webworker" />
/**
 * Web Worker: memformat DOCX di browser user. Berkas naskah tidak pernah diunggah ke server, jadi
 * tidak ada batas ukuran 4,5 MB; halaman tetap responsif selama proses.
 */
import { formatDocxBytes } from "../services/format-core";
import { browserAi } from "../services/ai-client";
import type { FormatConfig } from "../schemas/format-config";
import type { FrontFieldInput } from "../front/types";

export type WorkerRequest = { bytes: ArrayBuffer; fileName: string; config: FormatConfig; frontFields?: FrontFieldInput };
export type WorkerResponse = { ok: true; result: Awaited<ReturnType<typeof formatDocxBytes>> } | { ok: false; error: string };

self.onmessage = async (ev: MessageEvent<WorkerRequest>) => {
  const { bytes, fileName, config, frontFields } = ev.data;
  try {
    const result = await formatDocxBytes(new Uint8Array(bytes), fileName, config, frontFields, browserAi);
    (self as unknown as Worker).postMessage({ ok: true, result } satisfies WorkerResponse);
  } catch (e) {
    (self as unknown as Worker).postMessage({ ok: false, error: e instanceof Error ? e.message : String(e) } satisfies WorkerResponse);
  }
};