/**
 * Sampul template General: isi sampul dibaca dari dokumen (aturan + AI),
 * lalu sampul disusun ulang dengan tata letak baku:
 *
 *   JUDUL (+ terjemahan judul)                     ← paling atas
 *   ── ruang ──
 *   SKRIPSI / USULAN PENELITIAN                    ┐
 *   Diajukan untuk memenuhi …                      │
 *   Oleh: / NAMA / NIM / data mahasiswa lain       │ blok tengah
 *   Pembimbing                                     │
 *   [logo]                                         ┘
 *   ── ruang ──
 *   Prodi / Fakultas / Universitas / Kota / Tahun  ← frame di margin bawah
 *
 * Pengaman:
 *  - nilai dari AI wajib berupa salinan teks sampul asli,
 *  - data dari label eksplisit ("Nama : …", "NIM : …") diutamakan atas tebakan AI,
 *  - baris yang tidak terpetakan tetap dimasukkan (tidak ada teks hilang),
 *  - huruf miring / subskrip / superskrip di judul & keterangan dipertahankan,
 *  - judul/nama tidak ditemukan → tidak disusun ulang (kembali ke tata letak lama).
 */
import { escapeXml } from "./xml";
import { inlineOf } from "./figure-split";
import { flatten, parseBlocks, type DocModel, type ParaNode, type TableNode } from "./document-model";
import type { Outline } from "../structure/outline";
import type { FormatRules } from "../rules/format-rules";
import { startsNewPage } from "../structure/roles";
import { paragraphLine, type VerifyRegion } from "./verify";

export type CoverFields = {
  docType?: string | null;
  title?: string | null;
  purpose?: string | null;
  authorLabel?: string | null;
  name?: string | null;
  idLabel?: string | null;
  studentId?: string | null;
  program?: string | null;
  faculty?: string | null;
  university?: string | null;
  city?: string | null;
  year?: string | null;
};
export type CoverAI = (lines: string[]) => Promise<CoverFields | null>;

export type CoverTemplateResult = {
  built: boolean;
  source: "ai" | "rules" | null;
  reason: string | null;
  fields: CoverFields;
  regions: VerifyRegion[];
  /** Tabel sampul yang diganti sampul baru (untuk pemeriksaan integritas). */
  removed: { tables: number; rows: number; cells: number };
};

type Key = keyof CoverFields;
type CoverStyle = ReturnType<FormatRules["cover"]>;
const KEYS: Key[] = ["docType", "title", "purpose", "authorLabel", "name", "idLabel", "studentId",
  "program", "faculty", "university", "city", "year"];

/* ------------------------------------------------------------------ */
/* Pola                                                                 */
/* ------------------------------------------------------------------ */

const DOC_WORD = /^(skripsi|tesis|thesis|disertasi|dissertation|tugas\s+akhir|final\s+project|undergraduate\s+thesis|bachelor\s+thesis|proposal(\s+(skripsi|penelitian|tesis|tugas\s+akhir))?|usulan\s+penelitian(\s+skripsi)?|seminar\s+hasil(\s+penelitian)?|draf?t?\s+skripsi|laporan\s+[a-z\s]{3,30})$/i;
const DOC_CORE = /\b(skripsi|tesis|thesis|disertasi|dissertation|proposal|usulan|tugas\s+akhir|final\s+project|seminar|laporan|karya\s+(?:tulis\s+)?ilmiah)\b/i;
const DOC_VOCAB = new Set(("skripsi tesis thesis disertasi dissertation proposal usulan penelitian research tugas akhir " +
  "final project undergraduate bachelor master seminar hasil draf draft laporan karya tulis ilmiah sidang ujian").split(" "));

/** Baris yang HANYA berisi kata jenis dokumen ("USULAN PROPOSAL SKRIPSI", "(Proposal Penelitian)", "SKRIPSI / THESIS"). */
function isDocLine(l: string): boolean {
  const s = l.replace(/[():.\-–\/]+/g, " ").replace(/\s+/g, " ").trim();
  if (!s || s.length > 70) return false;
  const words = s.toLowerCase().split(" ");
  if (words.length <= 7 && words.every((w) => DOC_VOCAB.has(w)) && DOC_CORE.test(s)) return true;
  return l.split(/\s*\/\s*/).every((p) => DOC_WORD.test(p.replace(/[:.]$/, "").trim()));
}

const PURPOSE_RE = /^(?:diajukan|disusun|ditulis|dibuat)\s+(?!oleh\b)|\b(?:disusun|diajukan|ditulis|dibuat)\s+(?:untuk|sebagai|guna|dalam\s+rangka)\b|\b(?:untuk|guna)\s+memenuhi\b|sebagai\s+salah\s+satu\s+syarat|memperoleh\s+gelar|sebagian\s+persyaratan|sebagai\s+bahan\s+seminar|submitted\s+(?:in|to|as)\b|partial\s+fulfil/i;
const PURPOSE_CONT = /gelar|sarjana|\(S\.|program\s+studi|jurusan|submitted|partial\s+fulfil|requirement|degree|bachelor/i;
const EN_PURPOSE = /^(submitted|presented|in\s+partial|a\s+(thesis|proposal|dissertation)|an\s+undergraduate|as\s+(a|partial))\b/i;
const AUTHOR_RE = /^(disusun\s+oleh|diajukan\s+oleh|ditulis\s+oleh|dibuat\s+oleh|oleh|penyusun|by)\s*:?$/i;
const AUTHOR_HEAD = /^((?:disusun|diajukan|ditulis|dibuat)\s+oleh|oleh|penyusun)\s*[:：]\s*(?=\S)/i;
const NAME_LABEL = /^(nama(\s+mahasiswa)?|name|penulis)$/i;
const ID_LABEL = /^(nim|npm|nrp|nirm|nis|no\.?\s*mhs|nomor\s+induk(\s+mahasiswa)?|student\s+id)$/i;
const ID_RE = /\b(NIM|NPM|NRP|NIRM|NIS|No\.?\s*Mhs|Nomor\s+Induk(?:\s+Mahasiswa)?)\b\s*(?:\([^)]*\))?\s*[.:：]?\s*([0-9][0-9A-Za-z.\-\/ ]{3,22}[0-9A-Za-z])/i;
const TITLE_LABEL = /^(judul|title)(\s*\/\s*(judul|title))?$/i;
const PROGRAM_RE = /\b(program\s+studi|prodi|jurusan|departemen|department|program\s+sarjana|undergraduate\s+program|major)\b/i;
const FACULTY_RE = /\b(fakultas|faculty|school\s+of|business\s+school)\b/i;
const UNIV_RE = /\b(universitas|university|institut|institute|sekolah\s+tinggi|politeknik|akademi|stmik|stie|stkip|iain|uin)\b/i;
const INST_RE = new RegExp(`${PROGRAM_RE.source}|${FACULTY_RE.source}|${UNIV_RE.source}`, "i");
const MONTHS = /^(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember|january|february|march|may|june|july|august|october|december)$/i;
const LV_RE = /^([^:：]{2,45}?)\s*[:：]\s*(.+)$/;
const ADVISOR_RE = /^(?:(?:dosen|calon|komisi|tim)\s+)*(?:pembimbing|promotor|ko-?promotor|supervisors?|advisors?|co-?advisors?)\b/i;
const DEGREE_RE = /\b(?:Prof|Dr|Drs|Dra|Ir|H|Hj)\.|,\s*[SMD][a-z]{0,4}\.\s?[A-Za-z]|\bPh\.?\s?D\b|\bAkt\b/;
const NIP_RE = /^(NIP|NIDN|NIK|NUPTK)\b/i;
const ENUM = /^(\d+)\s*[.)]\s*/;
const STOPW = /\b(dan|di|ke|dari|pada|terhadap|yang|untuk|dengan|dalam|atas|sebagai|melalui|berbasis|studi|analisis|pengaruh|sistem|implementasi|perancangan|of|the|and|in|on|for|to|with|an?|by)\b/i;
const BINOMIAL = /^[A-Z][a-z]+(?:\s+[a-z]+){1,2}(?:\s+(?:L|var\.?\s*[a-z]+|subsp\.?\s*[a-z]+))?$|^[A-Z][a-z]+\s+spp?$/;
const LETTER = /[A-Za-z\u00C0-\u024F]/;

const norm = (s: string) => s.toLowerCase().replace(/\s+/g, " ").trim();
const capFirst = (s: string) => s.charAt(0).toUpperCase() + s.slice(1);
const stripParens = (s: string) => s.replace(/\s*\([^)]*\)/g, "").trim();
const idLabelOf = (l: string) =>
  /nomor\s+induk|student\s+id/i.test(l) ? "NIM" : stripParens(l).replace(/\s+/g, " ").toUpperCase();

function parseYear(line: string): { year: string; city?: string } | null {
  if (line.length > 50) return null;
  const c = line.replace(/(\d)\s+(?=\d)/g, "$1").trim();
  const m = /^(?:(.*?)[\s,]+)?((?:19|20)\d{2})$/.exec(c);
  if (!m) return null;
  let pre = (m[1] ?? "").replace(/[,\s]+$/, "").trim();
  if (/^tahun$/i.test(pre)) pre = "";
  if (pre && MONTHS.test(pre)) return { year: `${pre} ${m[2]}` };
  if (pre && (/\d/.test(pre) || INST_RE.test(pre) || pre.length > 30)) return null;
  return { year: m[2], city: pre || undefined };
}

const isNameLike = (t: string) =>
  t.length >= 3 && t.length <= 70 && !/\d/.test(t) && /^[A-Za-z\u00C0-\u024F .,'’\-]+$/.test(t) &&
  !isDocLine(t) && !AUTHOR_RE.test(t) && !INST_RE.test(t) && !LV_RE.test(t);

/** Nama orang tanpa label: 2–6 kata, tanpa kata sambung khas judul. */
const isPersonLike = (t: string) => {
  const w = t.split(/\s+/).length;
  return isNameLike(t) && w >= 2 && w <= 6 && t.length <= 50 && !STOPW.test(t);
};

const isAdvisorName = (l: string) => {
  if (l.length > 100 || INST_RE.test(l) || parseYear(l) || AUTHOR_RE.test(l) || ID_RE.test(l) || PURPOSE_RE.test(l)) return false;
  const core = l.replace(ENUM, "");
  return DEGREE_RE.test(core) || NIP_RE.test(core) || (ENUM.test(l) && core.length > 3) || isNameLike(core);
};

/* ------------------------------------------------------------------ */
/* Teks berformat (miring / subskrip / superskrip per karakter)         */
/* ------------------------------------------------------------------ */

type Rich = { t: string; f: number[] };
const IT = 1;
const SUB = 2;
const SUP = 4;

const plain = (t: string): Rich => ({ t, f: Array.from({ length: t.length }, () => 0) });
const slice = (r: Rich, s: number, e: number): Rich => ({ t: r.t.slice(s, e), f: r.f.slice(s, e) });

function trimRich(r: Rich): Rich {
  let s = 0;
  let e = r.t.length;
  while (s < e && /\s/.test(r.t[s])) s++;
  while (e > s && /\s/.test(r.t[e - 1])) e--;
  return slice(r, s, e);
}

function cutAt(r: Rich, points: number[]): Rich[] {
  const ps = [...new Set(points)].filter((p) => p > 0 && p < r.t.length).sort((a, b) => a - b);
  const out: Rich[] = [];
  let s = 0;
  for (const p of [...ps, r.t.length]) {
    const piece = trimRich(slice(r, s, p));
    if (piece.t) out.push(piece);
    s = p;
  }
  return out;
}

function cleanRich(r: Rich): Rich {
  let t = "";
  const f: number[] = [];
  for (let k = 0; k < r.t.length; k++) {
    const c = r.t[k];
    if (/\s/.test(c)) {
      if (t && !t.endsWith(" ")) { t += " "; f.push(r.f[k]); }
      continue;
    }
    t += c;
    f.push(r.f[k]);
  }
  let out = trimRich({ t, f });
  // Spasi di antara dua kata miring ikut miring ("Ananas comosus" satu potongan).
  for (let k = 1; k < out.t.length - 1; k++) {
    if (out.t[k] === " " && (out.f[k - 1] & IT) && (out.f[k + 1] & IT)) out.f[k] |= IT;
  }
  // "B A N D U N G" → "BANDUNG"
  if (/^(?:[A-Za-z\u00C0-\u024F] ){2,}[A-Za-z\u00C0-\u024F]$/.test(out.t)) {
    const keep = [...out.t].map((c, k) => (c === " " ? -1 : k)).filter((k) => k >= 0);
    out = { t: keep.map((k) => out.t[k]).join(""), f: keep.map((k) => out.f[k]) };
  }
  return out;
}

function joinRich(parts: Rich[]): { r: Rich; starts: number[] } {
  let t = "";
  const f: number[] = [];
  const starts: number[] = [];
  parts.forEach((p, k) => {
    if (k) { t += " "; f.push((f[f.length - 1] ?? 0) & (p.f[0] ?? 0) & IT); }
    starts.push(t.length);
    t += p.t;
    f.push(...p.f);
  });
  return { r: { t, f }, starts };
}

function italicShare(r: Rich): number {
  let letters = 0;
  let it = 0;
  for (let k = 0; k < r.t.length; k++) {
    if (!LETTER.test(r.t[k])) continue;
    letters++;
    if (r.f[k] & IT) it++;
  }
  return letters ? it / letters : 0;
}
const isItalic = (r: Rich) => italicShare(r) > 0.6;
const withoutItalic = (r: Rich): Rich => ({ t: r.t, f: r.f.map((x) => x & ~IT) });
const allItalic = (r: Rich): Rich => ({ t: r.t, f: r.f.map((x) => x | IT) });

const EN_W = /\b(the|of|and|in|on|for|to|with|a|an|by|from|at|using|based|effect|analysis|study|case|between|toward|towards)\b/gi;
const ID_W = /\b(dan|di|ke|dari|pada|terhadap|yang|untuk|dengan|dalam|atas|studi|kasus|analisis|pengaruh|antara|melalui|berbasis|perusahaan|serta|empiris)\b/gi;
/** Teks berbahasa Inggris (untuk mengenali terjemahan judul). */
const isEnglish = (t: string) => (t.match(EN_W)?.length ?? 0) > (t.match(ID_W)?.length ?? 0);

function fixEnum(r: Rich): Rich {
  const m = ENUM.exec(r.t);
  if (!m) return r;
  const pre = plain(`${m[1]}. `);
  const rest = slice(r, m[0].length, r.t.length);
  return { t: pre.t + rest.t, f: [...pre.f, ...rest.f] };
}

/** Nilai field dicari di gabungan baris-barisnya lalu dipecah per paragraf (format ikut terbawa). */
function groupsOf(rich: Rich[], idx: number[], value: string, brk: (k: number) => boolean): Rich[] {
  const v = value.replace(/\s+/g, " ").trim();
  if (!idx.length) return [plain(v)];
  const { r, starts } = joinRich(idx.map((i) => rich[i]));
  let s = r.t.indexOf(v);
  if (s < 0) s = r.t.toLowerCase().indexOf(v.toLowerCase());
  if (s < 0) return [plain(v)];
  const e = s + v.length;
  const cuts = starts.flatMap((st, k) => (k > 0 && st > s && st < e && brk(k) ? [st - s] : []));
  return cutAt(slice(r, s, e), cuts);
}

/* ------------------------------------------------------------------ */
/* Baris sampul (paragraf & tabel)                                      */
/* ------------------------------------------------------------------ */

const decode = (s: string) =>
  s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

const RUN_RE = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g;
const PIECE_RE = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\s*\/>|<w:br(?![^>]*w:type="page")[^>]*\/>|<w:cr\s*\/>/g;

function flagOf(rPr: string): number {
  let f = 0;
  if (/<w:i(?:\s+w:val="(?:1|true|on)")?\s*\/>/.test(rPr)) f |= IT;
  const va = /<w:vertAlign\s+w:val="(\w+)"/.exec(rPr)?.[1];
  if (va === "subscript") f |= SUB;
  else if (va === "superscript") f |= SUP;
  return f;
}

function linesOf(xml: string): Rich[] {
  const out: Rich[] = [];
  let cur: Rich = { t: "", f: [] };
  const add = (s: string, fl: number) => {
    cur.t += s;
    for (let k = 0; k < s.length; k++) cur.f.push(fl);
  };
  for (const run of xml.matchAll(RUN_RE)) {
    const inner = run[1];
    const fl = flagOf(/<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(inner)?.[1] ?? "");
    for (const m of inner.matchAll(PIECE_RE)) {
      if (m[1] !== undefined) add(decode(m[1]), fl);
      else if (m[0].startsWith("<w:tab")) add(" ", 0);
      else { out.push(cur); cur = { t: "", f: [] }; }
    }
  }
  out.push(cur);
  return out.map(cleanRich).filter((r) => r.t);
}

function drawingRuns(xml: string): string[] {
  return [...xml.matchAll(/<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g)]
    .map((m) => m[0])
    .filter((r) => /<w:drawing\b|<w:pict\b|<mc:AlternateContent\b/.test(r) && !/<w:txbxContent/.test(r));
}

/** "PROGRAM STUDI X FAKULTAS Y UNIVERSITAS Z" → tiga baris (judul & "Label : isi" tidak dipecah). */
const INST_HEAD = /^(program\s+studi|prodi|jurusan|departemen|fakultas|universitas|institut|sekolah\s+tinggi|politeknik)\b/i;
const INST_SPLIT = /(?=\b(?:program\s+studi|jurusan|departemen|fakultas|universitas|institut|sekolah\s+tinggi|politeknik)\b)/gi;

function instCuts(line: string): number[] {
  if (line.length > 160 || !INST_HEAD.test(line) || LV_RE.test(line) || PURPOSE_RE.test(line)) return [];
  return [...line.matchAll(INST_SPLIT)].map((m) => m.index ?? 0).filter((i) => i > 0);
}

/** Pecah baris gabungan: institusi, serta "Diajukan oleh:Nama : X" → "Diajukan oleh:" + "Nama : X". */
function splitRich(r: Rich): Rich[] {
  const out: Rich[] = [];
  for (const piece of cutAt(r, instCuts(r.t))) {
    const m = AUTHOR_HEAD.exec(piece.t);
    out.push(...(m ? cutAt(piece, [m[0].length]) : [piece]));
  }
  return out.filter((x) => x.t);
}

/* ------------------------------------------------------------------ */
/* Ekstraksi berbasis aturan                                            */
/* ------------------------------------------------------------------ */

type Extract = {
  f: CoverFields;
  at: Partial<Record<Key, number[]>>;
  sure: Set<Key>;
  dropped: Set<number>;
  advisors: number[];
};

function extractByRules(lines: string[], logoAt: number): Extract {
  const f: CoverFields = {};
  const at: Partial<Record<Key, number[]>> = {};
  const sure = new Set<Key>();
  const used = new Set<number>();
  const dropped = new Set<number>();
  const n = lines.length;
  const free = (i: number) => i >= 0 && i < n && !used.has(i);
  const set = (k: Key, v: string, idx: number[], certain = false) => {
    if (f[k]) return;
    f[k] = v; at[k] = idx; idx.forEach((i) => used.add(i));
    if (certain) sure.add(k);
  };

  // 0. Label kosong ("Judul / Title :")
  lines.forEach((l, i) => { if (TITLE_LABEL.test(l.replace(/\s*[:：]\s*$/, ""))) { used.add(i); dropped.add(i); } });

  // 1. Pasangan "Label : isi". Prodi/fakultas berlabel = data mahasiswa → tetap di blok identitas.
  lines.forEach((l, i) => {
    if (!free(i)) return;
    const m = LV_RE.exec(l);
    if (!m || /\d{1,2}[:：]\d{2}/.test(l)) return;
    const rawLabel = m[1].trim();
    if (ADVISOR_RE.test(rawLabel)) return;
    const label = stripParens(rawLabel).replace(/\s*\/.*$/, "").trim();
    const value = m[2].trim();
    if (NAME_LABEL.test(label) && isNameLike(value)) set("name", value, [i], true);
    else if (ID_LABEL.test(label) || ID_LABEL.test(stripParens(rawLabel))) {
      set("studentId", value.replace(/\s+/g, " "), [i], true);
      f.idLabel = idLabelOf(label); sure.add("idLabel");
    } else if (TITLE_LABEL.test(label)) set("title", value, [i], true);
  });

  // 1b. Pembimbing: "Pembimbing 1 : Dr. X" atau label lalu nama-nama di baris berikutnya.
  const advisors: number[] = [];
  for (let i = 0; i < n; i++) {
    if (!free(i) || !ADVISOR_RE.test(lines[i])) continue;
    const block = [i];
    if (!LV_RE.test(lines[i])) {
      for (let k = i + 1; k < n && free(k) && k !== logoAt && block.length < 8; k++) {
        if (ADVISOR_RE.test(lines[k]) || !isAdvisorName(lines[k])) break;
        block.push(k);
      }
    }
    for (const k of block) { used.add(k); advisors.push(k); }
    i = block[block.length - 1];
  }

  // 2. Kalimat keterangan (bisa beberapa baris, termasuk terjemahan Inggris)
  const pStart = lines.findIndex((l, i) => free(i) && PURPOSE_RE.test(l) && !AUTHOR_RE.test(l));
  let pEnd = -1;
  if (pStart !== -1) {
    pEnd = pStart;
    while (pEnd + 1 < n && pEnd + 1 !== logoAt && free(pEnd + 1)) {
      const next = lines[pEnd + 1];
      const cont = /^[a-z(]/.test(next) || PURPOSE_CONT.test(next) || PURPOSE_RE.test(next);
      if (!cont || AUTHOR_RE.test(next) || ID_RE.test(next) || parseYear(next) || UNIV_RE.test(next) || LV_RE.test(next) || next.length > 160) break;
      pEnd++;
    }
    const idx = [...Array(pEnd - pStart + 1).keys()].map((k) => pStart + k);
    set("purpose", idx.map((i) => lines[i]).join(" "), idx);
  }

  // 3. Tahun & kota (dari bawah)
  for (let i = n - 1; i >= 0; i--) {
    if (!free(i)) continue;
    const y = parseYear(lines[i]);
    if (!y) continue;
    set("year", y.year, [i]);
    if (y.city) { f.city = y.city; at.city = [i]; }
    else if (free(i - 1)) {
      const c = lines[i - 1];
      if (c.length <= 40 && !/\d/.test(c) && !INST_RE.test(c) && !LV_RE.test(c)) set("city", c, [i - 1]);
    }
    break;
  }

  // 4. Universitas, fakultas, program studi (baris tanpa label, dari bawah)
  const findLast = (re: RegExp) => {
    for (let i = n - 1; i >= 0; i--) if (free(i) && re.test(lines[i]) && !PURPOSE_RE.test(lines[i]) && !LV_RE.test(lines[i])) return i;
    return -1;
  };
  for (const [k, re] of [["university", UNIV_RE], ["faculty", FACULTY_RE], ["program", PROGRAM_RE]] as const) {
    const i = findLast(re);
    if (i !== -1) set(k, lines[i], [i]);
  }

  // 5. Jenis dokumen (semua baris jenis dokumen digabung), label penulis, NIM tanpa "Label : isi"
  const docIdx = lines.map((l, i) => (free(i) && isDocLine(l) ? i : -1)).filter((i) => i >= 0);
  if (docIdx.length) set("docType", docIdx.map((i) => lines[i].replace(/[:.]$/, "")).join(" "), docIdx, true);
  const a = lines.findIndex((l, i) => free(i) && AUTHOR_RE.test(l));
  if (a !== -1) set("authorLabel", capFirst(lines[a].replace(/\s*:?\s*$/, ":")), [a], true);
  if (!f.studentId) {
    const s = lines.findIndex((l, i) => free(i) && ID_RE.test(l));
    if (s !== -1) {
      const m = ID_RE.exec(lines[s])!;
      f.idLabel = idLabelOf(m[1]);
      set("studentId", m[2].trim(), [s], true);
      const pre = lines[s].slice(0, m.index).replace(/[\s,\/\-–(:]+$/, "").trim();
      if (pre && isNameLike(pre) && !f.name) { f.name = pre; at.name = [s]; sure.add("name"); }
    }
  }

  // 6. Nama tanpa label: setelah "Oleh", tepat di atas NIM, atau baris nama setelah jenis dokumen/keterangan.
  if (!f.name) {
    const after = a !== -1 ? [...Array(n).keys()].find((k) => k > a && free(k)) ?? -1 : -1;
    const sIdx = at.studentId?.[0] ?? -1;
    if (after !== -1 && isNameLike(lines[after])) set("name", lines[after], [after]);
    else if (sIdx > 0 && free(sIdx - 1) && isNameLike(lines[sIdx - 1])) set("name", lines[sIdx - 1], [sIdx - 1]);
  }
  if (!f.name) {
    const docMax = docIdx.length ? Math.max(...docIdx) : -1;
    const titleBeforeDoc = docMax > 0 && [...Array(docMax).keys()].some((i) => free(i));
    let anchor = Math.max(pEnd, titleBeforeDoc ? docMax : -1);
    if (anchor === -1 && logoAt > 0) anchor = logoAt - 1;
    if (anchor !== -1) {
      for (let i = anchor + 1; i < n; i++) {
        if (!free(i)) continue;
        if (INST_RE.test(lines[i]) || parseYear(lines[i])) break;
        if (isPersonLike(lines[i])) { set("name", lines[i], [i]); break; }
      }
    }
  }
  const nIdx = at.name?.[0] ?? -1;
  if (!f.studentId && nIdx !== -1 && free(nIdx + 1) && /^[0-9][0-9.\-\/]{5,19}$/.test(lines[nIdx + 1])) {
    set("studentId", lines[nIdx + 1], [nIdx + 1]);
  }

  // 7. Judul: baris bebas di bagian atas, sebelum keterangan/logo/penulis.
  //    Logo hanya jadi batas bila ada teks bebas di atasnya (logo paling atas tidak memotong judul).
  if (!f.title) {
    const logoStop = logoAt > 0 && [...Array(logoAt).keys()].some((i) => free(i)) ? logoAt : -1;
    const stops = [pStart, logoStop, a, nIdx, at.studentId?.[0] ?? -1].filter((x) => x >= 0);
    const stop = stops.length ? Math.min(...stops) : n;
    const idx: number[] = [];
    for (let i = 0; i < stop; i++) {
      if (!free(i) || LV_RE.test(lines[i])) continue;
      if (INST_RE.test(lines[i]) && idx.length === 0 && lines[i].length < 60) continue; // kop institusi di atas judul
      idx.push(i);
    }
    if (idx.length) set("title", idx.map((i) => lines[i]).join(" "), idx);
  }
  return { f, at, sure, dropped, advisors };
}

/* ------------------------------------------------------------------ */
/* Penggabungan AI + aturan — setiap nilai harus dipetakan ke baris asli */
/* ------------------------------------------------------------------ */

const coreOf = (l: string) => { const m = LV_RE.exec(l); return m ? m[2].trim() : l; };
const variants = (l: string) => [norm(l), norm(l.replace(/\s*[:：]\s*/g, " ")), norm(coreOf(l))];

/** Cari baris (atau rangkaian baris berurutan) yang PERSIS memuat nilai v. */
function locate(lines: string[], v: string, taken: Set<number>, multi: boolean): number[] | null {
  const t = norm(v);
  if (!t) return null;
  for (let i = 0; i < lines.length; i++) if (!taken.has(i) && variants(lines[i]).includes(t)) return [i];
  if (multi) {
    for (let i = 0; i < lines.length; i++) {
      if (taken.has(i)) continue;
      let acc = "";
      for (let j = i; j < lines.length && !taken.has(j); j++) {
        acc = acc ? `${acc} ${norm(lines[j])}` : norm(lines[j]);
        if (acc === t) return [...Array(j - i + 1).keys()].map((k) => i + k);
        if (!t.startsWith(acc)) break;
      }
    }
  }
  return null;
}

type Merged = { f: CoverFields; at: Map<Key, number[]>; fromAi: boolean };
const ORDER: Key[] = ["title", "purpose", "docType", "studentId", "name", "authorLabel", "year", "city", "university", "faculty", "program"];

function merge(lines: string[], rules: Extract, ai: CoverFields | null): Merged {
  const taken = new Set<number>([...rules.dropped, ...rules.advisors]);
  const f: CoverFields = {};
  const at = new Map<Key, number[]>();
  let fromAi = false;

  const accept = (k: Key, v: string, idx: number[]) => {
    f[k] = v;
    at.set(k, idx);
    // Kota & tahun boleh berbagi satu baris ("Jakarta 2026").
    if (k !== "year" && k !== "city") idx.forEach((i) => taken.add(i));
  };

  const fromAI = (k: Key): { v: string; idx: number[] } | null => {
    const raw = ai?.[k];
    if (typeof raw !== "string" || raw.trim().length < 2) return null;
    const v = raw.replace(/\s+/g, " ").trim();
    if (k === "year" || k === "city") {
      for (let i = lines.length - 1; i >= 0; i--) {
        const y = parseYear(lines[i]);
        if (k === "year" && y && norm(y.year) === norm(parseYear(v)?.year ?? v)) return { v: y.year, idx: [i] };
        if (k === "city" && ((y?.city && norm(y.city) === norm(v)) || (!taken.has(i) && norm(lines[i]) === norm(v)))) return { v, idx: [i] };
      }
      return null;
    }
    const idx = locate(lines, v, taken, k === "title" || k === "purpose");
    if (!idx) return null;
    const line = lines[idx[0]];
    const lv = LV_RE.exec(line);
    switch (k) {
      case "authorLabel": return AUTHOR_RE.test(line.replace(/\s*:?\s*$/, "")) ? { v: capFirst(line.replace(/\s*:?\s*$/, ":")), idx } : null;
      case "docType": {
        if (!isDocLine(line)) return null;
        // Baris jenis dokumen lain yang masih bebas ikut digabung.
        const more = lines.map((l, i) => (!taken.has(i) && !idx.includes(i) && isDocLine(l) ? i : -1)).filter((i) => i >= 0);
        const all = [...idx, ...more].sort((x, y) => x - y);
        return { v: all.map((i) => lines[i].replace(/[:.]$/, "")).join(" "), idx: all };
      }
      case "program": case "faculty": case "university":
        // Berlabel ("Prodi : S1 Akuntansi") = data mahasiswa, bukan blok institusi.
        return lv ? null : { v: line, idx };
      case "name": return { v: coreOf(line), idx };
      case "studentId": return { v: coreOf(line).replace(/^\s*(NIM|NPM|NRP)\b\s*[.:]?\s*/i, ""), idx };
      default: return { v: idx.map((i) => lines[i]).join(" "), idx };
    }
  };

  for (const k of ORDER) {
    const r = rules.f[k];
    const rIdx = rules.at[k];
    const ruleOk = !!r && !!rIdx && rIdx.every((i) => !taken.has(i) || k === "year" || k === "city");
    if (rules.sure.has(k) && ruleOk) { accept(k, r!, rIdx!); continue; }
    const a = fromAI(k);
    // AI hanya mengambil sebagian judul/keterangan yang ditemukan aturan → pakai versi lengkapnya.
    if (a && ruleOk && (k === "title" || k === "purpose") && rIdx!.length > a.idx.length && a.idx.every((i) => rIdx!.includes(i))) {
      accept(k, r!, rIdx!);
      continue;
    }
    if (a) { accept(k, a.v, a.idx); if (norm(a.v) !== norm(r ?? "")) fromAi = true; continue; }
    if (ruleOk) accept(k, r!, rIdx!);
  }
  const idl = ai?.idLabel;
  f.idLabel = rules.f.idLabel ?? (typeof idl === "string" && /^(nim|npm|nrp|nirm|nis)$/i.test(idl.trim()) ? idl.trim().toUpperCase() : null);
  return { f, at, fromAi };
}

/* ------------------------------------------------------------------ */
/* Pembagian baris ke blok                                              */
/* ------------------------------------------------------------------ */

function layoutBlocks(lines: string[], m: Merged, docType: string, dropped: Set<number>, advisors: number[]) {
  const f = m.f;
  const taken = new Set<number>([...dropped, ...advisors]);
  for (const idx of m.at.values()) idx.forEach((i) => taken.add(i));
  const finals = KEYS.filter((k) => k !== "idLabel" && f[k]).map((k) => norm(f[k]!)).concat(norm(docType));
  // Baris dianggap terwakili bila dipetakan ke field, atau isinya SAMA PERSIS dengan field.
  const consumed = (i: number) => taken.has(i) || variants(lines[i]).some((v) => finals.includes(v));

  const firstOf = (k: Key) => m.at.get(k)?.[0] ?? -1;
  const authorAt = [firstOf("authorLabel"), firstOf("name"), firstOf("studentId")].filter((x) => x >= 0).sort((a, b) => a - b)[0] ?? -1;

  const inst: { i: number; text: string }[] = [];
  for (const k of ["program", "faculty", "university"] as const) {
    if (f[k]) inst.push({ i: firstOf(k) === -1 ? 9_000 : firstOf(k), text: f[k]! });
  }
  const top: number[] = [];
  const identity: number[] = [];
  lines.forEach((line, i) => {
    if (consumed(i)) return;
    if (INST_RE.test(line) && !LV_RE.test(line) && !PURPOSE_RE.test(line)) inst.push({ i, text: line });
    else if (authorAt !== -1 && i > authorAt) identity.push(i);
    else top.push(i);
  });

  const seen = new Set<string>();
  const ordered = inst.sort((a, b) => a.i - b.i).map((x) => x.text)
    .filter((t) => !seen.has(norm(t)) && (seen.add(norm(t)), true));
  if (f.city) ordered.push(f.city);
  if (f.year) ordered.push(f.year);
  return { top, identity, inst: ordered };
}

/* ------------------------------------------------------------------ */
/* Logo                                                                 */
/* ------------------------------------------------------------------ */

function normalizeLogo(run: string, maxTw: number, minTw = 0): { xml: string; heightTw: number; widthTw: number } {
  // Logo melayang → inline (deklarasi namespace di elemen anchor ikut dibawa).
  let r = inlineOf(run);
  const ext = /<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/.exec(r);
  if (!ext) return { xml: r, heightTw: 2000, widthTw: 2000 };
  let cx = Number(ext[1]);
  let cy = Number(ext[2]);
  const big = Math.max(cx, cy);
  const maxEmu = maxTw * 635;
  const minEmu = Math.min(minTw, maxTw) * 635;
  // Terlalu besar → diperkecil; terlalu kecil → diperbesar ke ukuran standar sampul.
  const target = big > maxEmu ? maxEmu : big < minEmu ? minEmu : big;
  if (target !== big && big > 0) {
    const k = target / big;
    cx = Math.round(cx * k);
    cy = Math.round(cy * k);
    r = r.replace(/(<wp:extent\s+)cx="\d+"\s+cy="\d+"/, `$1cx="${cx}" cy="${cy}"`)
      .replace(/(<a:ext\s+)cx="\d+"\s+cy="\d+"/, `$1cx="${cx}" cy="${cy}"`);
  }
  return { xml: r, heightTw: Math.round(cy / 635), widthTw: Math.round(cx / 635) };
}

/* ------------------------------------------------------------------ */
/* Penyusunan                                                           */
/* ------------------------------------------------------------------ */

const PG_BORDERS = /<w:pgBorders\b[^>]*\/>|<w:pgBorders\b[\s\S]*?<\/w:pgBorders>/g;

type Group = (ParaNode | TableNode)[];
type BuildOpts = { style: CoverStyle; font: string; width: number; height: number; docTypeFallback: string; ai?: CoverAI };

/**
 * Satu dokumen bisa punya beberapa halaman sampul (sampul muka + sampul dalam, sampul
 * Indonesia + Inggris). Tiap halaman (dipisah section/page break) disusun sendiri.
 */
export async function buildCoverFromTemplate(
  model: DocModel,
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  opts: BuildOpts,
): Promise<CoverTemplateResult> {
  const none = { tables: 0, rows: 0, cells: 0 };
  const fail = (reason: string): CoverTemplateResult => ({ built: false, source: null, reason, fields: {}, regions: [], removed: none });
  const body = model.body;

  const coverNodes = nodes.filter((_n, i) => outline.items[i]?.zone === "cover");
  if (!coverNodes.length) return fail("halaman sampul tidak terdeteksi.");
  const pos = coverNodes.map((n) => body.indexOf(n as DocModel["body"][number]));
  if (pos.some((p) => p < 0) || pos.some((p, k) => k > 0 && p !== pos[k - 1] + 1)) {
    return fail("bagian sampul tidak berurutan dalam dokumen.");
  }

  const pages: Group[] = [[]];
  for (const n of coverNodes) {
    pages[pages.length - 1].push(n);
    if (n.kind === "paragraph" && (n.hasSectPr || n.hasPageBreak)) pages.push([]);
  }
  const hasText = (g: Group) => g.some((n) => n.kind === "table" || Boolean(n.text.trim()) || n.hasDrawing);
  const groups: Group[] = [];
  for (const g of pages) {
    if (!g.length) continue;
    if (hasText(g) || !groups.length) groups.push(g);
    else groups[groups.length - 1].push(...g);
  }
  if (groups.length > 1 && !hasText(groups[0])) groups[1].unshift(...groups.shift()!);

  // Dari halaman terakhir ke pertama, supaya posisi halaman sebelumnya di body tidak bergeser.
  const results: CoverTemplateResult[] = [];
  for (let g = groups.length - 1; g >= 0; g--) results.unshift(await buildPage(model, nodes, outline, groups[g], opts));
  const ok = results.filter((r) => r.built);
  if (!ok.length) return results[0];
  return {
    built: true,
    source: ok.some((r) => r.source === "ai") ? "ai" : "rules",
    reason: results.find((r) => !r.built)?.reason ?? null,
    fields: ok[0].fields,
    regions: ok.flatMap((r) => r.regions),
    removed: ok.reduce((a, r) => ({
      tables: a.tables + r.removed.tables, rows: a.rows + r.removed.rows, cells: a.cells + r.removed.cells,
    }), none),
  };
}

async function buildPage(
  model: DocModel,
  nodes: (ParaNode | TableNode)[],
  outline: Outline,
  coverNodes: Group,
  opts: BuildOpts,
): Promise<CoverTemplateResult> {
  const none = { tables: 0, rows: 0, cells: 0 };
  const fail = (reason: string): CoverTemplateResult => ({ built: false, source: null, reason, fields: {}, regions: [], removed: none });
  const body = model.body;
  const pos = coverNodes.map((n) => body.indexOf(n as DocModel["body"][number]));
  if (pos.some((p) => p < 0)) return fail("bagian sampul tidak berurutan dalam dokumen.");

  /* ---- baris teks berformat, logo, section break ---- */
  const rich: Rich[] = [];
  const gap: boolean[] = [];          // baris didahului paragraf kosong / logo / tabel
  const logos: string[] = [];
  let logoAt = -1;
  let sectPr: string | null = null;
  let pendingGap = false;
  const removed = { tables: 0, rows: 0, cells: 0 };
  const beforeLines: string[] = [];

  const push = (rs: Rich[]) => {
    rs.forEach((r, k) => { rich.push(r); gap.push(k === 0 && pendingGap); });
    if (rs.length) pendingGap = false;
  };
  const noteLogo = (p: ParaNode) => {
    const runs = drawingRuns(p.xml);
    if (!runs.length) return false;
    if (logoAt === -1) logoAt = rich.length;
    logos.push(...runs);
    return true;
  };
  const read = (p: ParaNode): Rich[] => {
    beforeLines.push(paragraphLine(p.xml));
    return linesOf(p.xml).flatMap(splitRich);
  };

  type Cell = TableNode["rows"][number]["cells"][number];
  /** Isi sel: paragraf + tabel bersarang, urut sesuai dokumen. */
  const readCell = (cell: Cell): { lines: Rich[]; nested: boolean } => {
    const out: Rich[] = [];
    let nested = false;
    const take = (x: ParaNode | TableNode) => {
      if (x.kind === "paragraph") { noteLogo(x); out.push(...read(x)); }
      else { nested = true; out.push(...readTable(x)); }
    };
    for (const ch of cell.children) {
      if (ch.kind === "paragraph" || ch.kind === "table") take(ch);
      else if (ch.kind === "container") flatten(ch.children).forEach(take);
    }
    return { lines: out, nested };
  };
  /** Baris "Label | : isi" digabung jadi satu baris; sel berisi blok (beberapa baris / tabel) dibaca per sel. */
  const readTable = (t: TableNode): Rich[] => {
    removed.tables++;
    const out: Rich[] = [];
    for (const row of t.rows) {
      removed.rows++;
      const cells: { lines: Rich[]; nested: boolean }[] = [];
      for (const cell of row.cells) {
        removed.cells++;
        const c = readCell(cell);
        if (c.lines.length) cells.push(c);
      }
      const simple = cells.length > 1 && cells.length <= 4 && cells.every((c) => !c.nested && c.lines.length === 1);
      if (simple) out.push(...splitRich(cleanRich(joinRich(cells.map((c) => c.lines[0])).r)));
      else for (const c of cells) out.push(...c.lines);
    }
    return out;
  };

  for (const n of coverNodes) {
    if (n.kind === "table") {
      push(readTable(n));
      pendingGap = true;
      continue;
    }
    const hasLogo = noteLogo(n);
    if (hasLogo) pendingGap = true;
    const rs = read(n);
    push(rs);
    if (!rs.length || hasLogo) pendingGap = true;
    if (n.hasSectPr) sectPr = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(n.xml)?.[0] ?? sectPr;
  }
  const lines = rich.map((r) => r.t);
  if (!lines.length) return fail("sampul tidak berisi teks.");

  /* ---- isi sampul ---- */
  const rules = extractByRules(lines, logoAt);
  let ai: CoverFields | null = null;
  if (opts.ai) { try { ai = await opts.ai(lines); } catch { ai = null; } }
  const merged = merge(lines, rules, ai);
  const f = merged.f;
  if (!f.title) return fail("judul pada sampul tidak dapat dikenali.");
  if (!f.name) return fail("nama penulis pada sampul tidak dapat dikenali.");

  const docType = f.docType ??
    ({ skripsi: "SKRIPSI", tesis: "TESIS", proposal: "PROPOSAL SKRIPSI" } as Record<string, string>)[opts.docTypeFallback] ?? "SKRIPSI";
  const blocks = layoutBlocks(lines, merged, docType, rules.dropped, rules.advisors);

  // Judul: dipecah per paragraf asli yang dipisah baris kosong, terjemahan dalam kurung, atau baris miring.
  const tIdx = merged.at.get("title") ?? [];
  const titleGroups = groupsOf(rich, tIdx, f.title, (k) => {
    const line = lines[tIdx[k]];
    if (gap[tIdx[k]]) return true;
    return isEnglish(line) && (/^\(.*\)$/.test(line) || (isItalic(rich[tIdx[k]]) && !isItalic(rich[tIdx[k - 1]])));
  }).map((r, k, all) => {
    // Terjemahan = kelompok berbahasa Inggris di bawah judul yang bukan bahasa Inggris.
    const tr = k > 0 && isEnglish(r.t) && !isEnglish(all[0].t);
    if (tr) return { r: italicShare(r) === 0 ? allItalic(r) : r, tr };
    return { r: italicShare(r) > 0.95 ? withoutItalic(r) : r, tr };
  });

  // Keterangan: baris lanjutan kalimat digabung; kalimat baru (mis. terjemahan Inggris) jadi paragraf sendiri.
  const pIdx = merged.at.get("purpose") ?? [];
  const purposeGroups = f.purpose
    ? groupsOf(rich, pIdx, f.purpose, (k) => {
      const prev = rich[pIdx[k - 1]];
      const next = rich[pIdx[k]];
      if (!/^[A-Z]/.test(next.t)) return false;
      return EN_PURPOSE.test(next.t) || /[.!?]$/.test(prev.t) || isItalic(next) !== isItalic(prev);
    })
    : [];

  // Jenis dokumen: selalu 12 pt. Tahap ("USULAN PENELITIAN") dan jenis karya ("SKRIPSI") di dua baris.
  const DOC_HALF = 24;
  const FINAL_RE = /\b(skripsi|tesis|thesis|disertasi|dissertation|tugas\s+akhir|final\s+project|karya\s+(?:tulis\s+)?ilmiah)\b/i;
  const STAGE_RE = /\b(usulan|proposal|seminar|draf|draft|laporan|hasil|penelitian|research)\b/i;
  const dm = /^(.*\S)\s*(\([^)]*\))$/.exec(docType);
  const docMain = dm ? dm[1] : docType;
  const docNote = dm ? dm[2] : null;
  const docLines = ((): string[] => {
    if (docMain.includes("/")) return [docMain];            // "SKRIPSI / THESIS" tetap satu baris
    const m = FINAL_RE.exec(docMain);
    if (!m) return [docMain];
    const before = docMain.slice(0, m.index).replace(/[\s\-–:,]+$/, "").trim();
    const after = docMain.slice(m.index + m[0].length).replace(/^[\s\-–:,]+/, "").trim();
    if (before && !after && STAGE_RE.test(before)) return [before, m[0]];
    if (!before && after && STAGE_RE.test(after)) return [after, m[0]]; // "SKRIPSI USULAN PENELITIAN" → dibalik
    return [docMain];
  })();

  const idLine = f.studentId ? `${f.idLabel ?? "NIM"}. ${f.studentId}` : null;
  const advisors = [...rules.advisors].sort((x, y) => x - y)
    .map((i) => ({ r: fixEnum(rich[i]), label: ADVISOR_RE.test(lines[i]) && !LV_RE.test(lines[i]) }));

  /* ---- ukuran & jarak ---- */
  const st = opts.style;
  const W = opts.width;
  const H = opts.height;
  const tHalf = st.titleHalf;
  const xHalf = st.textHalf;
  const trHalf = st.translationSame ? tHalf : Math.max(xHalf, tHalf - 4);
  // "Oleh:" hanya bila sampul asli memilikinya, atau pedoman memintanya.
  const useLabel = Boolean(f.authorLabel) || st.authorLabel;
  const lineTw = (half: number) => Math.round(half * 10 * 1.15);
  const h = (t: string, half: number, wide: boolean) =>
    Math.max(1, Math.ceil((t.length * half * 10 * (wide ? 0.62 : 0.52)) / W)) * lineTw(half);
  const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);

  const titleH = sum(titleGroups.map((g) => h(g.r.t, g.tr ? trHalf : tHalf, !g.tr) + (g.tr ? 120 : 0)));
  const docH = sum(docLines.map((t) => h(t, DOC_HALF, true))) + (docNote ? h(docNote, DOC_HALF, false) : 0);
  const purposeH = purposeGroups.length ? 240 + sum(purposeGroups.map((g) => h(g.t, xHalf, false))) : 0;
  const topH = blocks.top.length ? 120 + sum(blocks.top.map((i) => h(lines[i], xHalf, false))) : 0;
  const authorH = (useLabel ? lineTw(xHalf) : 0) + h(f.name, xHalf, true) + (idLine ? lineTw(xHalf) : 0) +
    sum(blocks.identity.map((i) => h(lines[i], xHalf, false)));
  const advH = advisors.length
    ? 240 + sum(advisors.map((a) => h(a.r.t, xHalf, false))) + 120 * Math.max(0, advisors.filter((a) => a.label).length - 1)
    : 0;
  const instH = sum(blocks.inst.map((t) => h(t, st.institutionHalf, true)));
  const fixed = titleH + docH + purposeH + topH + authorH + advH + instH;
  const SECT_LINE = 60;

  // Empat ruang kosong: judul→jenis dokumen, keterangan→"Oleh", identitas→logo, logo→institusi.
  const hasLogo = logos.length > 0;
  const weight = { top: 1, author: 0.7, logo: hasLogo ? 1 : 0, bottom: 1 };
  const minGap = { top: 360, author: 240, logo: hasLogo ? 360 : 0, bottom: 480 };
  const minSum = minGap.top + minGap.author + minGap.logo + minGap.bottom;
  const wSum = weight.top + weight.author + weight.logo + weight.bottom;

  const logoXml: string[] = [];
  let logoH = 0;
  let maxLogo = Math.round(st.logoMaxTwips * 1.25);
  const minLogo = Math.round(maxLogo * 0.85);
  for (let pass = 0; pass < 3; pass++) {
    logoXml.length = 0;
    logoH = 0;
    let dim = 0;
    for (const r of logos) {
      const l = normalizeLogo(r, maxLogo, Math.min(minLogo, maxLogo));
      logoXml.push(l.xml);
      logoH = Math.max(logoH, l.heightTw);
      dim = Math.max(dim, l.heightTw, l.widthTw);
    }
    const deficit = minSum - (H - fixed - logoH - SECT_LINE);
    if (deficit <= 0 || !hasLogo || maxLogo <= 900) break;
    maxLogo = Math.max(900, Math.floor(dim * Math.max(0.2, (logoH - deficit) / Math.max(1, logoH))));
  }

  const room = Math.max(0, H - fixed - logoH - SECT_LINE);
  const k = room < minSum ? room / minSum : 1;
  const extra = Math.max(0, room - minSum);
  const gapOf = (key: keyof typeof weight) => Math.round(minGap[key] * k + (extra * weight[key]) / wSum);
  const gapTop = gapOf("top");
  const gapAuthor = gapOf("author");
  const gapLogo = gapOf("logo");

  /* ---- XML ---- */
  const fn = escapeXml(opts.font);
  const rPr = (half: number, bold: boolean, fl = 0) =>
    `<w:rPr><w:rFonts w:ascii="${fn}" w:hAnsi="${fn}" w:cs="${fn}" w:eastAsia="${fn}"/>` +
    (bold ? "<w:b/><w:bCs/>" : `<w:b w:val="0"/><w:bCs w:val="0"/>`) +
    (fl & IT ? "<w:i/><w:iCs/>" : `<w:i w:val="0"/><w:iCs w:val="0"/>`) +
    `<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>` +
    (fl & SUB ? `<w:vertAlign w:val="subscript"/>` : fl & SUP ? `<w:vertAlign w:val="superscript"/>` : "") +
    `</w:rPr>`;
  const runsXml = (r: Rich, half: number, bold: boolean, caps: boolean) => {
    let out = "";
    for (let a = 0; a < r.t.length;) {
      let b = a + 1;
      while (b < r.t.length && r.f[b] === r.f[a]) b++;
      const fl = r.f[a];
      let seg = r.t.slice(a, b);
      // Nama ilmiah miring ("Ananas comosus") tidak ikut dikapitalkan.
      if (caps && !((fl & IT) && BINOMIAL.test(seg.replace(/^[\s(]+|[\s.,;:)]+$/g, "")))) seg = seg.toUpperCase();
      out += `<w:r>${rPr(half, bold, fl)}<w:t xml:space="preserve">${escapeXml(seg)}</w:t></w:r>`;
      a = b;
    }
    return out;
  };
  const frame = `<w:framePr w:w="${W}" w:wrap="around" w:vAnchor="margin" w:hAnchor="margin" w:xAlign="center" w:yAlign="bottom"/>`;
  type PO = { half: number; bold?: boolean; caps?: boolean; before?: number; frame?: boolean };
  const paraRich = (r: Rich, o: PO) =>
    `<w:p><w:pPr>${o.frame ? frame : ""}` +
    `<w:spacing w:before="${Math.max(0, Math.round(o.before ?? 0))}" w:after="0" w:line="240" w:lineRule="auto"/>` +
    `<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="center"/>${rPr(o.half, !!o.bold)}</w:pPr>` +
    `${runsXml(r, o.half, !!o.bold, !!o.caps)}</w:p>`;
  const para = (t: string, o: PO) => paraRich(plain(t), o);

  const out: string[] = [];
  // Judul (paling atas)
  titleGroups.forEach((g) => out.push(g.tr
    ? paraRich(g.r, { half: trHalf, bold: true, before: 120 })
    : paraRich(g.r, { half: tHalf, bold: true, caps: st.titleCaps })));
  // Blok tengah: jenis dokumen (12 pt), keterangan, identitas, pembimbing, logo
  docLines.forEach((t, i) => out.push(para(t, { half: DOC_HALF, bold: true, caps: true, before: i === 0 ? gapTop : 0 })));
  if (docNote) out.push(para(docNote, { half: DOC_HALF }));
  purposeGroups.forEach((g, i) => out.push(paraRich(g, { half: xHalf, before: i === 0 ? 240 : 0 })));
  blocks.top.forEach((li, i) => out.push(paraRich(rich[li], { half: xHalf, before: i === 0 ? 120 : 0 })));
  if (useLabel) out.push(para(f.authorLabel ?? "Oleh:", { half: xHalf, before: gapAuthor }));
  out.push(para(f.name, { half: xHalf, bold: true, caps: st.nameCaps, before: useLabel ? 0 : gapAuthor }));
  if (idLine) out.push(para(idLine, { half: xHalf, bold: st.idBold }));
  for (const li of blocks.identity) out.push(paraRich(rich[li], { half: xHalf }));
  advisors.forEach((a, i) => out.push(paraRich(a.r, { half: xHalf, before: i === 0 ? 240 : a.label ? 120 : 0 })));
  if (logoXml.length) {
    out.push(`<w:p><w:pPr><w:spacing w:before="${gapLogo}" w:after="0" w:line="240" w:lineRule="auto"/>` +
      `<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr>${logoXml.join("")}</w:p>`);
  }
  // Identitas institusi (frame di margin bawah)
  for (const t of blocks.inst) out.push(para(t, { half: st.institutionHalf, bold: st.institutionBold, caps: st.institutionCaps, frame: true }));

  // Paragraf penutup dibuat sangat kecil agar tidak mendorong sampul ke halaman kedua.
  const tiny = `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>`;
  const after = body[pos[pos.length - 1] + 1];
  const afterIdx = after ? nodes.indexOf(after as ParaNode | TableNode) : -1;
  const afterOpens = afterIdx !== -1 && startsNewPage(outline.items[afterIdx]?.role ?? "body");
  if (sectPr) out.push(`<w:p><w:pPr>${tiny}${sectPr.replace(PG_BORDERS, "")}</w:pPr></w:p>`);
  else if (after && !afterOpens) out.push(`<w:p><w:pPr>${tiny}</w:pPr><w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:br w:type="page"/></w:r></w:p>`);

  const startId = 2_000_000 + body.reduce((m, n) => (n.kind === "paragraph" || n.kind === "table" ? Math.max(m, n.id) : m), 0);
  const fresh = parseBlocks(out.join(""), startId) as ParaNode[];
  body.splice(pos[0], pos.length, ...fresh);

  const text = fresh.map((p) => p.text).join(" ");
  return {
    built: true,
    source: merged.fromAi ? "ai" : "rules",
    reason: null,
    fields: { ...f, docType },
    regions: [{ beforeLines: beforeLines.filter(Boolean), afterLines: fresh.map((p) => paragraphLine(p.xml)).filter(Boolean), beforeText: text, afterText: text }],
    removed,
  };
}