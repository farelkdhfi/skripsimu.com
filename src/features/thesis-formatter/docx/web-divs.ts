/**
 * Kotak HTML tersembunyi (`w:divId` → `word/webSettings.xml`).
 *
 * Teks yang ditempel dari halaman web / pengelola referensi (Mendeley, Zotero) sering membawa
 * "div" HTML: paragrafnya menunjuk ke div yang punya margin (dan kadang garis) sendiri. Word
 * MENAMBAHKAN margin div itu ke indentasi paragraf — LibreOffice mengabaikannya. Penulis biasanya
 * menutupinya dengan indentasi minus; setelah formatter menulis indentasi baku, margin div
 * membuat teks bergeser (mis. Daftar Pustaka 2,54 cm ke kanan).
 *
 * Paragraf yang indentasinya DITULIS formatter dan menunjuk ke div bermargin/bergaris dilepas dari
 * div-nya, jadi indentasi yang ditulis = yang tampil. Paragraf yang indentasinya tidak diubah
 * (masih sama dengan naskah asli) tetap menunjuk ke div-nya supaya tampilannya tetap seperti asli.
 * Yang dibuang hanya penunjuk tata letak itu; teks tidak berubah. Dijalankan pada semua bagian yang diformat
 * (lampiran beku dikembalikan utuh setelahnya, lengkap dengan div aslinya).
 */
import type JSZip from "jszip";
import type { BlockNode, DocModel, ParaNode } from "./document-model";

const indOf = (xml: string) => /<w:ind\b[^>]*\/>/.exec(/<w:pPr>[\s\S]*?<\/w:pPr>/.exec(xml)?.[0] ?? "")?.[0] ?? "";

/** Catat indentasi asli paragraf yang menunjuk ke div. Panggil pada model ASLI. */
export function divIndents(model: DocModel): Map<ParaNode, string> {
    const out = new Map<ParaNode, string>();
    const walk = (list: BlockNode[]) => {
        for (const n of list) {
            if (n.kind === "paragraph") { if (/<w:pPr>[\s\S]*?<w:divId\b[\s\S]*?<\/w:pPr>/.test(n.xml)) out.set(n, indOf(n.xml)); }
            else if (n.kind === "container") walk(n.children);
            else if (n.kind === "table") for (const r of n.rows) for (const c of r.cells) walk(c.children);
        }
    };
    walk(model.body);
    return out;
}

type Div = { parent: string | null; left: number; right: number; top: number; bottom: number; border: boolean };

function readDivs(xml: string): Map<string, Div> {
    const out = new Map<string, Div>();
    // Telusuri tag secara berurutan untuk mengetahui induk tiap div (div bersarang lewat w:divsChild).
    const stack: string[] = [];
    const re = /<w:div\s+w:id="(-?\d+)"\s*>|<\/w:div>|<w:(marLeft|marRight|marTop|marBottom)\s+w:val="(-?\d+)"\s*\/>|<w:(top|left|bottom|right)\b[^>]*w:val="([^"]+)"[^>]*\/>/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml))) {
        if (m[1] !== undefined) {
            out.set(m[1], { parent: stack[stack.length - 1] ?? null, left: 0, right: 0, top: 0, bottom: 0, border: false });
            stack.push(m[1]);
        } else if (m[0] === "</w:div>") stack.pop();
        else {
            const d = out.get(stack[stack.length - 1] ?? "");
            if (!d) continue;
            if (m[2] === "marLeft") d.left = Number(m[3]);
            else if (m[2] === "marRight") d.right = Number(m[3]);
            else if (m[2] === "marTop") d.top = Number(m[3]);
            else if (m[2] === "marBottom") d.bottom = Number(m[3]);
            else if (m[5] && m[5] !== "none" && m[5] !== "nil") d.border = true;
        }
    }
    return out;
}

/** Div (beserta induknya) menggeser/menghiasi isi: ada margin atau garis. */
function effective(divs: Map<string, Div>, id: string): boolean {
    for (let d = divs.get(id), k = 0; d && k < 20; d = d.parent ? divs.get(d.parent) : undefined, k++) {
        if (d.left || d.right || d.top || d.bottom || d.border) return true;
    }
    return false;
}

export async function detachWebDivs(zip: JSZip, model: DocModel, original: Map<ParaNode, string>): Promise<number> {
    const ws = await zip.file("word/webSettings.xml")?.async("string");
    if (!ws || !/<w:divs>/.test(ws)) return 0;
    const divs = readDivs(ws);
    let count = 0;
    const strip = (xml: string) => xml.replace(/<w:divId\s+w:val="(-?\d+)"\s*\/>/g, (all, id: string) => {
        if (!effective(divs, id)) return all;
        count++;
        return "";
    });
    const walk = (list: BlockNode[]) => {
        for (const n of list) {
            if (n.kind === "paragraph") {
                if (!n.xml.includes("<w:divId")) continue;
                // Indentasi tidak diubah formatter → div tetap (tampilan asli).
                if (original.has(n) && original.get(n) === indOf(n.xml)) continue;
                n.xml = strip(n.xml);
            }
            else if (n.kind === "container") walk(n.children);
            else if (n.kind === "table") {
                // Baris tabel: div-nya dibiarkan (posisi tabel diatur tblInd/jc, bukan di sini).
                for (const r of n.rows) for (const c of r.cells) walk(c.children);
            }
        }
    };
    walk(model.body);
    return count;
}