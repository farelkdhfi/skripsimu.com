import { NextResponse } from "next/server";
import { ZodError } from "zod";

export function errorResponse(e: unknown) {
  if (e instanceof ZodError)
    return NextResponse.json({ error: "Input tidak valid: " + e.issues.map((i) => i.message).join(", ") }, { status: 400 });
  const status = e instanceof Error && "status" in e && typeof e.status === "number" ? e.status : 500;
  if (status >= 500) console.error(e);
  return NextResponse.json({ error: e instanceof Error ? e.message : "Terjadi kesalahan" }, { status });
}