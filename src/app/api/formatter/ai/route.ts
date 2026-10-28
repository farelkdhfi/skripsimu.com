/**
 * Endpoint AI untuk pemformatan di browser. Hanya menerima TEKS kecil (daftar paragraf terpotong,
 * baris sampul, daftar baris bagian depan) — berkas naskah tidak pernah dikirim ke sini.
 */
import { NextResponse } from "next/server";
import { classifyNodes } from "@/features/thesis-formatter/services/analyze.service";
import { extractCoverAI } from "@/features/thesis-formatter/services/cover-ai.service";
import { extractFrontAI } from "@/features/thesis-formatter/services/front-ai.service";
import type { Signals } from "@/features/thesis-formatter/structure/roles";

export const maxDuration = 300;

const MAX_BODY = 4_000_000;

export async function POST(req: Request) {
  const raw = await req.text();
  if (raw.length > MAX_BODY) return NextResponse.json({ error: "Data terlalu besar." }, { status: 413 });
  let body: { kind?: string; payload?: unknown };
  try { body = JSON.parse(raw); } catch { return NextResponse.json({ error: "JSON tidak sah." }, { status: 400 }); }

  switch (body.kind) {
    case "classify": {
      if (!Array.isArray(body.payload)) return NextResponse.json({ error: "payload harus array." }, { status: 400 });
      const signals = (body.payload as Signals[]).map((s) => ({
        text: String(s?.text ?? "").slice(0, 600), hasDrawing: Boolean(s?.hasDrawing), isTable: Boolean(s?.isTable),
      }));
      return NextResponse.json(await classifyNodes(signals));
    }
    case "cover": {
      const lines = Array.isArray(body.payload) ? (body.payload as unknown[]).map(String).slice(0, 200) : [];
      return NextResponse.json(await extractCoverAI(lines));
    }
    case "front": {
      const listing = typeof body.payload === "string" ? body.payload.slice(0, 60_000) : "";
      return NextResponse.json(await extractFrontAI(listing));
    }
    default:
      return NextResponse.json({ error: "kind tidak dikenal." }, { status: 400 });
  }
}