/**
 * Pengambilan data dari halaman bagian depan. Setiap nilai adalah SALINAN teks naskah
 * (tidak dikarang). Baris yang tidak terpetakan ke data atau teks baku dicatat supaya
 * dilaporkan ke pengguna, bukan hilang diam-diam.
 */
import type { ParaNode } from "../docx/document-model";
import { richText } from "./read";
import {
    ABSTRACT_RE, ABSTRAK_RE, APPROVAL_RE, DOC_TYPE_RE, INST_RE, KP_RE, OLEH_RE, PENGESAHAN_HEAD_RE,
    PERNYATAAN_RE, YEAR_RE, bare, norm,
} from "./segment";
import type { FrontFields, FrontPage, Item, LineItem, Pic, PicItem, Rich } from "./types";

const isLine = (x: Item): x is LineItem => x.k === "line";
const isPic = (x: Item): x is PicItem => x.k === "pic";

export const emptyFields = (): FrontFields => ({
    judul: null, judulInggris: null, nama: null, npm: null, tahun: null,
    pembimbingUtama: null, pembimbingAnggota: null, ketuaSidang: null, dekan: null,
    tanggalSidang: null, tempatTanggalPernyataan: null, tempatTanggalKataPengantar: null,
});

/* ------------------------------------------------------------------ */
/* Bahasa & teks                                                        */
/* ------------------------------------------------------------------ */

const EN_W = /\b(the|of|and|in|on|for|to|with|a|an|by|from|at|using|based|effect|effects|analysis|study|case|between|toward|towards|peels?|rations?|fluid)\b/gi;
const ID_W = /\b(dan|di|ke|dari|pada|terhadap|yang|untuk|dengan|dalam|atas|studi|kasus|analisis|pengaruh|antara|melalui|berbasis|serta|kulit|buah|ransum|cairan|sebagai)\b/gi;
export const isEnglish = (t: string) => (t.match(EN_W)?.length ?? 0) > (t.match(ID_W)?.length ?? 0);

const italicShare = (r: Rich) => {
    let all = 0;
    let it = 0;
    for (const s of r) {
        const n = s.t.replace(/[^A-Za-z\u00C0-\u024F]/g, "").length;
        all += n;
        if (s.it) it += n;
    }
    return all ? it / all : 0;
};

/** Gabung beberapa baris menjadi satu teks berformat (tanda hubung di ujung baris tidak diberi spasi). */
export function joinRich(lines: Rich[]): Rich {
    const out: Rich = [];
    lines.forEach((r, k) => {
        if (k > 0) {
            const prev = out[out.length - 1];
            const glue = prev && /-$/.test(prev.t) ? "" : " ";
            if (glue) out.push({ t: glue, it: Boolean(prev?.it && r[0]?.it), sub: false, sup: false, b: false });
        }
        out.push(...r.map((s) => ({ ...s })));
    });
    return out;
}

export const plainRich = (t: string): Rich => [{ t, it: false, sub: false, sup: false, b: false }];

/**
 * Teks yang berisi dirinya sendiri dua kali ("Judul X Judul X", mis. judul terbaca dari dua tempat)
 * → satu kali. Pembanding: huruf & angka saja, tanpa huruf besar/kecil.
 */
export function collapseRepeat(r: Rich): Rich {
    const text = r.map((s) => s.t).join("");
    const at: number[] = [];
    let norm = "";
    [...text].forEach((ch, i) => { const c = ch.toLowerCase(); if (/[a-z0-9]/.test(c)) { norm += c; at.push(i); } });
    const n = norm.length;
    if (n < 8 || n % 2 || norm.slice(0, n / 2) !== norm.slice(n / 2)) return r;
    let left = at[n / 2];
    const out: Rich = [];
    for (const s of r) {
        if (left <= 0) break;
        out.push({ ...s, t: s.t.slice(0, left) });
        left -= s.t.length;
    }
    if (out.length) out[out.length - 1] = { ...out[out.length - 1], t: out[out.length - 1].t.replace(/\s+$/, "") };
    return out.filter((s) => s.t);
}

const DEGREE_RE = /\b(?:Prof|Dr|Drs|Dra|Ir|H|Hj)\.|,\s*[A-Z][A-Za-z]{0,5}\.|\bPh\.?\s?D\b/;
const LABELISH = /[:\uFF1A]\s*$/;
const AUTHOR_LABEL_RE = /^((disusun|diajukan|ditulis|dibuat)\s+oleh|oleh|by|penulis|penyusun|nama)$/i;
/** Nama orang (penulis): huruf, 1–6 kata, tanpa angka & label ("Disusun oleh:" bukan nama). */
export const isPersonName = (t: string) => {
    if (/[:\uFF1A]\s*$/.test(t.trim())) return false;
    const s = bare(t);
    return s.length >= 3 && s.length <= 70 && !/\d/.test(s) && /^[A-Za-z\u00C0-\u024F .,'\u2019-]+$/.test(s) &&
        s.split(" ").length <= 6 && !INST_RE.test(s) && !DOC_TYPE_RE.test(s) && !OLEH_RE.test(s) && !AUTHOR_LABEL_RE.test(s) &&
        !APPROVAL_RE.test(s);
};
/** Nama penanda tangan (boleh bergelar). */
const isSignerName = (t: string) => {
    const s = norm(t);
    return s.length >= 3 && s.length <= 120 && /[A-Za-z]/.test(s) && !LABELISH.test(s) && !/^(a\.?\s?n\.?|atas\s+nama)\b/i.test(s) &&
        !APPROVAL_RE.test(s) && !/^(nip|nidn|nik)\b/i.test(s) && (DEGREE_RE.test(s) || isPersonName(s));
};

const NPM_RE = /^(?:(?:npm|nim)\s*\.?\s*:?\s*)?(\d[\d .]{7,20}\d)$/i;
export const npmOf = (t: string) => {
    const m = NPM_RE.exec(norm(t));
    return m ? m[1].replace(/[\s.]/g, "") : null;
};

/* ------------------------------------------------------------------ */
/* Laporan pembacaan                                                    */
/* ------------------------------------------------------------------ */

export type Extraction = {
    fields: FrontFields;
    /** Baris naskah yang tidak terpakai (dilaporkan ke pengguna). */
    unmapped: { page: string; text: string }[];
    /** Teks baku naskah yang diganti teks pedoman, bila berbeda. */
    replaced: { page: string; from: string; to: string }[];
    /** Data yang tertulis berbeda di halaman lain (mis. NPM di pernyataan ≠ NPM di sampul). */
    conflicts: string[];
    signatures: Partial<Record<"pembimbingUtama" | "pembimbingAnggota" | "ketuaSidang" | "dekan", Pic>>;
    pernyataanPics: Pic[];
    kataPengantar?: ProsaPage & { closingDate: string | null; closingRole: string | null; closingPics: Pic[] };
    abstrak?: AbstractPage;
    abstract?: AbstractPage;
    /** Gambar/bentuk yang tidak dipakai lagi (bingkai buatan penulis, gambar tanpa pemilik). */
    dropped: { pic: Pic; why: string }[];
};

export type ProsaPage = { paras: ParaNode[] };
export type AbstractPage = ProsaPage & { title: Rich | null; name: string | null; keywords: ParaNode | null; after: ParaNode[] };

const simple = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

/** Isi field bila kosong; bila sudah ada dan berbeda, catat perbedaannya (nilai pertama dipakai). */
function setField(f: FrontFields, x: Extraction, key: "nama" | "npm", value: string, where: string) {
    const v = value.trim();
    if (!v) return;
    const cur = f[key];
    if (cur === null) { f[key] = v; return; }
    const same = key === "npm" ? cur.replace(/\D/g, "") === v.replace(/\D/g, "") : simple(cur) === simple(v);
    if (!same) x.conflicts.push(`${key === "npm" ? "NPM" : "Nama"} di ${where} ("${v}") berbeda dengan yang dipakai ("${cur}"). Periksa naskah.`);
}

/* ------------------------------------------------------------------ */
/* Sampul                                                               */
/* ------------------------------------------------------------------ */

function readCover(page: FrontPage, f: FrontFields, x: Extraction, spec: { institution: string[]; docType: string; purpose: string }) {
    const label = page.kind === "sampulMuka" ? "Sampul muka" : "Sampul kedua";
    const lines = page.items.filter(isLine);
    const used = new Set<LineItem>();
    const docAt = lines.findIndex((l) => DOC_TYPE_RE.test(bare(l.text)));

    // Judul: semua baris sebelum jenis dokumen. Terjemahan Inggris = baris miring/berbahasa Inggris di bawahnya.
    if (docAt > 0) {
        const titleLines = lines.slice(0, docAt).filter((l) => !(INST_RE.test(l.text) && l.text.length < 60));
        const en = (l: LineItem) => italicShare(l.segs) > 0.6 || isEnglish(l.text);
        let split = -1;
        for (let k = 1; k < titleLines.length; k++) {
            if (!en(titleLines[k - 1]) && titleLines.slice(k).every(en)) { split = k; break; }
        }
        const idLines = split === -1 ? titleLines : titleLines.slice(0, split);
        const enLines = split === -1 ? [] : titleLines.slice(split);
        if (!f.judul && idLines.length) f.judul = joinRich(idLines.map((l) => l.segs));
        if (!f.judulInggris && enLines.length) f.judulInggris = joinRich(enLines.map((l) => l.segs));
        titleLines.forEach((l) => used.add(l));
    }
    if (docAt !== -1) {
        used.add(lines[docAt]);
        const t = bare(lines[docAt].text);
        if (simple(t) !== simple(spec.docType)) x.replaced.push({ page: label, from: t, to: spec.docType });
    }

    // Sesudah jenis dokumen: kalimat tujuan, nama, NPM. Tanpa baris jenis dokumen, posisi nama tidak
    // pasti (judul pendek bisa mirip nama) → nama diambil dari halaman lain.
    let named = false;
    for (const l of docAt === -1 ? [] : lines.slice(docAt + 1)) {
        const t = bare(l.text);
        if (/^(diajukan|submitted)\b/i.test(t)) {
            used.add(l);
            if (simple(t) !== simple(spec.purpose)) x.replaced.push({ page: label, from: t, to: spec.purpose });
            continue;
        }
        const npm = npmOf(t);
        if (npm) { used.add(l); setField(f, x, "npm", npm, label.toLowerCase()); continue; }
        if (INST_RE.test(t) || YEAR_RE.test(t)) break;
        if (!named && isPersonName(l.text)) { setField(f, x, "nama", t, label.toLowerCase()); used.add(l); named = true; }
    }

    // Tahun (dari bawah). Lokasi kampus & institusi baku dari template.
    for (let k = lines.length - 1; k >= 0; k--) {
        const t = bare(lines[k].text);
        const m = /^(?:([A-Za-z][A-Za-z .'-]{1,30}),?\s+)?((?:19|20)\d{2})$/.exec(t);
        if (!m) continue;
        used.add(lines[k]);
        if (!f.tahun) f.tahun = m[2];
        if (!m[1] && k > 0 && !used.has(lines[k - 1])) {
            const c = bare(lines[k - 1].text);
            if (c.length <= 40 && !/\d/.test(c) && !INST_RE.test(c)) used.add(lines[k - 1]);
        }
        break;
    }
    const inst = lines.filter((l) => !used.has(l) && INST_RE.test(l.text));
    inst.forEach((l) => used.add(l));
    const want = spec.institution.map(simple);
    for (const l of inst) {
        if (!want.includes(simple(l.text))) x.replaced.push({ page: label, from: bare(l.text), to: spec.institution.join(" / ") });
    }

    for (const l of lines) if (!used.has(l)) x.unmapped.push({ page: label, text: l.text });

    // Logo & hiasan sampul dari template; gambar sampul naskah tidak dipakai.
    for (const p of page.items.filter(isPic).map((it) => it.pic)) {
        x.dropped.push({ pic: p, why: p.picture ? `logo/gambar di ${label.toLowerCase()} naskah (diganti logo template)` : `bentuk/kotak hiasan di ${label.toLowerCase()}` });
    }
}

/* ------------------------------------------------------------------ */
/* Lembar pengesahan                                                    */
/* ------------------------------------------------------------------ */

type SignerKey = "pembimbingUtama" | "pembimbingAnggota" | "ketuaSidang" | "dekan";

/** Label tanggal yang memang tanggal sidang/ujian sarjana (bukan seminar usulan/kolokium/dll.). */
export function isSidangDateLabel(label: string): boolean {
    return /\b(sidang|ujian)\b/i.test(label) && !/\b(sup|seminar|usulan|kolokium|proposal|up|hasil)\b/i.test(label);
}
/** Jabatan yang sama dengan label template (tidak perlu dilaporkan). */
const STANDARD_ROLE: Record<SignerKey, RegExp> = {
    pembimbingUtama: /^pembimbing\s+(utama|i|1|pertama|ketua)\b/i,
    pembimbingAnggota: /^pembimbing\s+(anggota|ii|2|kedua|pendamping)\b/i,
    ketuaSidang: /^ketua\s+(panitia\s+)?sidang\b/i,
    dekan: /^dekan\b/i,
};
const ROLE_NAME: Record<SignerKey, string> = {
    pembimbingUtama: "Pembimbing Utama", pembimbingAnggota: "Pembimbing Anggota",
    ketuaSidang: "Ketua Panitia Sidang", dekan: "Dekan",
};
/** Awalan atas nama/mewakili di depan jabatan: "a.n.", "An.", "u.b.", "Plh.", "Plt.". */
const ON_BEHALF = String.raw`(?:(?:a\.?\s?n|u\.?\s?b|plh|plt)\.?\s+)?`;
const roleRe = (body: string) => new RegExp(`^${ON_BEHALF}${body}`, "i");
const ROLE_OF: [SignerKey, RegExp][] = [
    ["pembimbingUtama", roleRe(String.raw`pembimbing\s+(utama|i|1|pertama|ketua)\b`)],
    ["pembimbingAnggota", roleRe(String.raw`pembimbing\s+(anggota|ii|2|kedua|pendamping)\b`)],
    // Penanda tangan ke-3 di bawah "Mengesahkan" yang bukan dekan: ketua panitia sidang/ penguji,
    // atau (sesuai kebiasaan fakultas) ketua program studi/ departemen/ jurusan/ komisi.
    ["ketuaSidang", roleRe(String.raw`ketua\s+(panitia|sidang|penguji|program\s+studi|prodi|departemen|jurusan|komisi)\b`)],
    // Wakil dekan yang menandatangani atas nama dekan → posisi tanda tangan dekan di template.
    ["dekan", roleRe(String.raw`(?:wakil\s+)?dekan\b`)],
];

function readPengesahan(page: FrontPage, f: FrontFields, x: Extraction) {
    const label = "Lembar pengesahan";
    const items = page.items;
    const lines = items.filter(isLine);
    const used = new Set<LineItem>();
    const olehAt = lines.findIndex((l) => OLEH_RE.test(bare(l.text)));
    lines.slice(0, Math.max(0, olehAt)).forEach((l) => used.add(l));          // judul (dipakai dari sampul)
    if (!f.judul && olehAt > 0) f.judul = joinRich(lines.slice(0, olehAt).map((l) => l.segs));
    if (olehAt !== -1) {
        used.add(lines[olehAt]);
        for (const l of lines.slice(olehAt + 1, olehAt + 4)) {
            const t = bare(l.text);
            const npm = npmOf(t);
            if (npm) { used.add(l); setField(f, x, "npm", npm, "lembar pengesahan"); continue; }
            if (isPersonName(l.text)) { used.add(l); setField(f, x, "nama", t, "lembar pengesahan"); }
        }
    }

    // Penanda tangan: nama di atas jabatannya ("Nama ↵ Pembimbing Utama") ATAU di bawahnya
    // ("Pembimbing Utama, ↵ [tanda tangan] ↵ Nama"). Susunan yang memberi pasangan terbanyak
    // dipakai untuk seluruh halaman; pencarian tidak melewati judul blok (Menyetujui/Mengesahkan)
    // atau jabatan lain, dan nama penulis tidak pernah dianggap penanda tangan.
    const isRole = (t: string) => ROLE_OF.some(([, re]) => re.test(bare(t)));
    const author = simple(f.nama ?? "");
    const roleAt: { i: number; key: SignerKey }[] = [];
    items.forEach((it, i) => {
        if (!isLine(it)) return;
        const role = ROLE_OF.find(([, re]) => re.test(bare(it.text)));
        if (role) { roleAt.push({ i, key: role[0] }); used.add(it); }
    });
    const roleOfPair = new Map<string, number>();       // "arah:jabatan" → indeks baris jabatan yang dipasangkan
    const pair = (dir: 1 | -1) => {
        const out = new Map<SignerKey, number>();
        const taken = new Set<number>();
        for (const r of roleAt) {
            for (let k = r.i + dir, seen = 0; k >= 0 && k < items.length && seen < 3; k += dir) {
                const c = items[k];
                if (!isLine(c)) continue;
                seen++;
                if (isRole(c.text) || APPROVAL_RE.test(bare(c.text)) || OLEH_RE.test(bare(c.text))) break;
                if (isSignerName(c.text) && simple(c.text) !== author && !taken.has(k) && !out.has(r.key)) {
                    out.set(r.key, k); taken.add(k); roleOfPair.set(`${dir}:${r.key}`, r.i); break;
                }
            }
        }
        return out;
    };
    const above = pair(-1);
    const below = pair(1);
    const signerAt = below.size > above.size ? below : above;     // indeks item baris nama
    const dirOf = new Map<SignerKey, 1 | -1>([...signerAt.keys()].map((k) => [k, signerAt === below ? 1 : -1]));
    // Susunan campuran (mis. "An. Ketua Program Studi ↵ Nama" sementara yang lain "Nama ↵
    // Jabatan"): jabatan yang belum berpasangan boleh memakai arah lain, asal barisnya belum dipakai.
    const other = signerAt === above ? below : above;
    const takenAt = new Set(signerAt.values());
    for (const [key, k] of other) {
        if (!signerAt.has(key) && !takenAt.has(k)) { signerAt.set(key, k); takenAt.add(k); dirOf.set(key, signerAt === above ? 1 : -1); }
    }
    for (const [key, k] of signerAt) {
        const c = items[k] as LineItem;
        used.add(c);
        if (!f[key]) f[key] = norm(c.text);
        // Jabatan di naskah berbeda dari label template ("An. Ketua Program Studi", "Wakil Dekan"):
        // nama tetap dipasang di posisi template (label template), tapi dilaporkan.
        const ri = roleOfPair.get(`${dirOf.get(key)}:${key}`);
        const roleText = ri !== undefined ? bare((items[ri] as LineItem).text) : "";
        if (roleText && !STANDARD_ROLE[key].test(roleText)) {
            x.conflicts.push(`Jabatan di lembar pengesahan naskah "${roleText}" (${norm(c.text)}) dipasang pada posisi ${ROLE_NAME[key]} di template; label jabatan mengikuti template. Periksa apakah sudah sesuai.`);
        }
    }
    for (const l of lines) {
        const t = bare(l.text);
        if (APPROVAL_RE.test(t) && t.length <= 20) used.add(l);                // "Menyetujui:", "Mengesahkan:"
        if (PENGESAHAN_HEAD_RE.test(t)) used.add(l);
        const d = /^tanggal\b([^:\uFF1A]*)[:\uFF1A]\s*(.*)$/i.exec(norm(l.text));
        if (d) {
            used.add(l);
            const v = d[2].replace(/[.\u2026:\s]+$/, "").trim();
            // Hanya tanggal SIDANG/UJIAN SARJANA yang mengisi slot template; tanggal seminar usulan
            // (SUP/kolokium) atau pelaksanaan lain bukan tanggal sidang → dikosongkan & dilaporkan.
            if (isSidangDateLabel(d[1])) { if (v && !f.tanggalSidang) f.tanggalSidang = v; }
            else if (v) x.conflicts.push(`Tanggal di lembar pengesahan naskah ("${norm(l.text).slice(0, 80)}") bukan tanggal sidang/ujian sarjana, jadi tidak dipakai; tanggal sidang di template dikosongkan.`);
        }
    }
    for (const l of lines) if (!used.has(l)) x.unmapped.push({ page: label, text: l.text });

    // Tanda tangan: milik nama penanda tangan TERDEKAT di bawah gambar (tanda tangan selalu di atas
    // nama), atau nama di paragraf yang sama bila gambar ditambatkan ke baris nama itu sendiri.
    const order = [...signerAt.entries()].sort((a, b) => a[1] - b[1]);
    items.forEach((it, i) => {
        if (!isPic(it)) return;
        if (!it.pic.picture) { x.dropped.push({ pic: it.pic, why: "bentuk/kotak di lembar pengesahan" }); return; }
        const sameLine = order.find(([, at]) => at < i && (items[at] as LineItem).para === it.para);
        const owner = sameLine ?? order.find(([, at]) => at > i);
        if (owner && !x.signatures[owner[0]]) x.signatures[owner[0]] = it.pic;
        else x.dropped.push({ pic: it.pic, why: "tanda tangan di lembar pengesahan yang pemiliknya tidak bisa dipastikan" });
    });
}

/* ------------------------------------------------------------------ */
/* Lembar pernyataan                                                    */
/* ------------------------------------------------------------------ */

function readPernyataan(page: FrontPage, f: FrontFields, x: Extraction, spec: { body: string[]; intro: string }) {
    const label = "Lembar pernyataan";
    const lines = page.items.filter(isLine);
    const used = new Set<LineItem>();
    const LV = /^([^:\uFF1A]{2,40}?)\s*[:\uFF1A]\s*(.*)$/;
    let after = false;
    let statement = false;      // sedang di dalam paragraf "Menyatakan …"/"Demikian …"
    for (let i = 0; i < lines.length; i++) {
        const l = lines[i];
        let t = norm(l.raw.replace(/\t/g, " "));
        // Label dan isinya di baris terpisah (sel tabel/paragraf sendiri): "Nama" ↵ ": Budi".
        const next = lines[i + 1];
        if (next && /^(nama|npm|nim|judul(\s+(skripsi|penelitian|usulan\s+penelitian))?)$/i.test(t) && /^[:\uFF1A]/.test(norm(next.raw))) {
            used.add(next);
            t = `${t} ${norm(next.raw.replace(/\t/g, " "))}`;
            i++;
        }
        // Lanjutan kalimat pernyataan yang terpotong Enter (diawali huruf kecil) ikut terpakai.
        if (statement && /^[a-z]/.test(t)) { used.add(l); continue; }
        statement = false;
        if (PERNYATAAN_RE.test(bare(t)) && t.length <= 120) { used.add(l); continue; }
        if (/^yang\s+bertanda\s+tangan/i.test(t)) {
            used.add(l);
            if (simple(t) !== simple(spec.intro)) x.replaced.push({ page: label, from: t, to: spec.intro });
            continue;
        }
        const lv = LV.exec(t);
        if (lv && /^nama\b/i.test(lv[1])) { used.add(l); setField(f, x, "nama", lv[2], "lembar pernyataan"); continue; }
        if (lv && /^(npm|nim)\b/i.test(lv[1])) { used.add(l); setField(f, x, "npm", npmOf(lv[2]) ?? lv[2], "lembar pernyataan"); continue; }
        if (lv && /^judul/i.test(lv[1])) { used.add(l); continue; }
        if (/^(menyatakan|demikian)\b/i.test(t)) {
            used.add(l);
            statement = true;
            if (!spec.body.some((b) => simple(b) === simple(t))) {
                const to = spec.body.find((b) => b.slice(0, 8).toLowerCase() === t.slice(0, 8).toLowerCase()) ?? spec.body.join(" ");
                x.replaced.push({ page: label, from: t, to });
            }
            continue;
        }
        const d = /^dibuat\s+di\s+(.+)$/i.exec(t);
        if (d) { used.add(l); if (!f.tempatTanggalPernyataan) f.tempatTanggalPernyataan = d[1].replace(/[.\u2026\s]+$/, "").trim(); continue; }
        if (/^(penulis|yang\s+menyatakan)\s*,?$/i.test(t)) { used.add(l); after = true; continue; }
        if (after && isPersonName(t.replace(/[()]/g, ""))) { used.add(l); setField(f, x, "nama", t.replace(/[()]/g, ""), "lembar pernyataan (di bawah tanda tangan)"); continue; }
    }
    // Lanjutan judul pada baris berikutnya ("Judul Skripsi : …" yang terpotong Enter) ikut terpakai.
    let inTitle = false;
    for (const l of lines) {
        const t = norm(l.raw.replace(/\t/g, " "));
        if (/^judul\b/i.test(t)) { inTitle = true; continue; }
        if (inTitle && !used.has(l) && !/^(menyatakan|demikian)\b/i.test(t)) { used.add(l); continue; }
        inTitle = false;
    }
    for (const l of lines) if (!used.has(l)) x.unmapped.push({ page: label, text: l.text });
    for (const p of page.items.filter(isPic)) {
        if (p.pic.picture) x.pernyataanPics.push(p.pic);
        else x.dropped.push({ pic: p.pic, why: "kotak bingkai buatan penulis di lembar pernyataan (diganti bingkai template)" });
    }
}

/* ------------------------------------------------------------------ */
/* Halaman prosa                                                        */
/* ------------------------------------------------------------------ */

/** Paragraf berurutan di halaman (satu paragraf bisa memuat beberapa baris Shift+Enter). */
function parasOf(items: Item[]): { p: ParaNode; text: string; pics: Pic[]; inTable: boolean }[] {
    const out: { p: ParaNode; text: string; pics: Pic[]; inTable: boolean }[] = [];
    for (const it of items) {
        if (it.k === "break") continue;
        const last = out[out.length - 1];
        const same = last && last.p === it.para;
        const rec = same ? last : { p: it.para, text: "", pics: [] as Pic[], inTable: it.k === "line" ? it.inTable : false };
        if (!same) out.push(rec);
        if (it.k === "line") rec.text = norm(`${rec.text} ${it.text}`);
        if (it.k === "pic") rec.pics.push(it.pic);
    }
    return out;
}

const DATE_LINE_RE = /^[A-Z][A-Za-z.'\- ]{1,40},\s*(?:\d{1,2}\s+)?[A-Za-z]{3,12}\s+\d{4}\.?$/;

function readKataPengantar(page: FrontPage, x: Extraction) {
    const label = "Kata pengantar";
    const ps = parasOf(page.items);
    const head = ps.findIndex((p) => KP_RE.test(bare(p.text)));
    let end = ps.length;
    let closingRole: string | null = null;
    let closingDate: string | null = null;
    const closingPics: Pic[] = [];
    const roleAt = (() => {
        for (let k = ps.length - 1; k > head; k--) {
            const t = bare(ps[k].text);
            if (!t && !ps[k].pics.length) continue;
            return /^(penulis|peneliti|hormat\s+saya|penyusun)\b/i.test(t) && t.length <= 30 ? k : -1;
        }
        return -1;
    })();
    if (roleAt !== -1) {
        closingRole = bare(ps[roleAt].text);
        end = roleAt;
        for (let k = roleAt - 1; k > head && k >= roleAt - 5; k--) {
            const t = norm(ps[k].text);
            if (DATE_LINE_RE.test(t)) { closingDate = t; end = k; break; }
        }
        for (let k = end; k <= roleAt; k++) {
            closingPics.push(...ps[k].pics.filter((p) => p.picture));
            for (const p of ps[k].pics) if (!p.picture) x.dropped.push({ pic: p, why: "bentuk di penutup kata pengantar" });
            const t = norm(ps[k].text);
            if (t && k !== roleAt && t !== closingDate) x.unmapped.push({ page: label, text: t });
        }
    }
    const body = ps.slice(head + 1, end).filter((p) => p.text).map((p) => p.p);
    for (const p of ps.slice(head + 1, end)) for (const pic of p.pics) x.dropped.push({ pic, why: "gambar di isi kata pengantar" });
    x.kataPengantar = { paras: body, closingDate, closingRole, closingPics };
    if (closingDate && !x.fields.tempatTanggalKataPengantar) x.fields.tempatTanggalKataPengantar = closingDate.replace(/\.$/, "");
}

function readAbstract(page: FrontPage, x: Extraction, english: boolean) {
    const label = english ? "Abstract" : "Abstrak";
    const ps = parasOf(page.items).filter((p) => p.text || p.pics.length);
    const head = ps.findIndex((p) => (english ? ABSTRACT_RE : ABSTRAK_RE).test(bare(p.text)));
    const before = ps.slice(0, Math.max(0, head));
    let name: string | null = null;
    let titleParas = before;
    if (before.length && isPersonName(before[before.length - 1].text)) {
        name = bare(before[before.length - 1].text);
        titleParas = before.slice(0, -1);
    }
    const lines = page.items.filter(isLine);
    const titleLines = lines.filter((l) => titleParas.some((p) => p.p === l.para));
    const title = titleLines.length ? joinRich(titleLines.map((l) => l.segs)) : null;
    const kwAt = ps.findIndex((p, k) => k > head && /^(kata\s*kunci|key\s*words?)\s*[:\uFF1A]/i.test(p.text));
    const bodyEnd = kwAt === -1 ? ps.length : kwAt;
    for (const p of ps) for (const pic of p.pics) x.dropped.push({ pic, why: `gambar/bentuk di halaman ${label.toLowerCase()}` });
    const rec: AbstractPage = {
        title, name,
        paras: ps.slice(head + 1, bodyEnd).filter((p) => p.text).map((p) => p.p),
        keywords: kwAt === -1 ? null : ps[kwAt].p,
        after: kwAt === -1 ? [] : ps.slice(kwAt + 1).map((p) => p.p),
    };
    if (english) x.abstract = rec;
    else x.abstrak = rec;
}

/* ------------------------------------------------------------------ */

export function extractFront(
    pages: FrontPage[],
    spec: { institution: string[]; docType: string; purpose: string; pernyataan: { body: string[]; intro: string } },
): Extraction {
    const f = emptyFields();
    const x: Extraction = {
        fields: f, unmapped: [], replaced: [], conflicts: [], signatures: {}, pernyataanPics: [], dropped: [],
    };
    // Sampul kedua lebih lengkap (NPM), jadi dibaca lebih dulu.
    const ordered = [...pages].sort((a, b) => Number(b.kind === "sampulKedua") - Number(a.kind === "sampulKedua"));
    for (const p of ordered) {
        if (p.kind === "sampulMuka" || p.kind === "sampulKedua") readCover(p, f, x, spec);
    }
    for (const p of pages) {
        if (p.kind === "pengesahan") readPengesahan(p, f, x);
        else if (p.kind === "pernyataan") readPernyataan(p, f, x, spec.pernyataan);
        else if (p.kind === "kataPengantar") readKataPengantar(p, x);
        else if (p.kind === "abstrak") readAbstract(p, x, false);
        else if (p.kind === "abstract") readAbstract(p, x, true);
    }
    return x;
}

export { richText };