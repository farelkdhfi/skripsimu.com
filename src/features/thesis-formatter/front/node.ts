/**
 * Khusus Node (server, skrip tes): template dibaca dari folder proyek dan rekaman diagnosis
 * disimpan sebagai berkas JSON. JANGAN diimpor dari kode browser.
 *
 * Panggil `useNodeIO()` sekali sebelum memformat di Node.
 */
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { setTemplateLoader, TEMPLATE_DIR } from "./template";
import { setDiagnosisSink } from "./index";

let done = false;

export function useNodeIO(): void {
    if (done) return;
    done = true;
    setTemplateLoader(async (file) => new Uint8Array(await readFile(path.join(process.cwd(), ...TEMPLATE_DIR, file))));
    setDiagnosisSink(async (payload) => {
        const dir = process.env.FRONT_DEBUG_DIR ?? (process.env.NODE_ENV === "production" ? null : path.join(process.cwd(), "front-debug"));
        if (!dir) {
            console.info("[front-template] rekaman diagnosis:", JSON.stringify(payload));
            return;
        }
        await mkdir(dir, { recursive: true });
        const name = `${new Date().toISOString().replace(/[:.]/g, "-")}-${String(payload.nama ?? "naskah").replace(/[^A-Za-z0-9]+/g, "_").slice(0, 40)}.json`;
        await writeFile(path.join(dir, name), JSON.stringify(payload, null, 2), "utf8");
        console.info(`[front-template] rekaman diagnosis: ${path.join(dir, name)}`);
    });
}