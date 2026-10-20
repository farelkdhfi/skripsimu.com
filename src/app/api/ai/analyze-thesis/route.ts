import { NextResponse } from "next/server";
import { errorResponse } from "@/lib/api-response";
import { parseUpload } from "@/features/file-upload/services/parse-upload";
import { extractParagraphs } from "@/features/file-upload/services/extract-text";
import { classifyParagraphs } from "@/features/thesis-formatter/services/analyze.service";

export const maxDuration = 120;

export async function POST(req: Request) {
  try {
    const { file } = await parseUpload(req);
    const { paragraphs, kind } = await extractParagraphs(file);
    const { roles, recommendations } = await classifyParagraphs(paragraphs);
    const headings = paragraphs.map((text, i) => ({ text, role: roles[i] })).filter((x) => x.role && x.role !== "body" && x.role !== "skip");
    return NextResponse.json({ kind, totalParagraphs: paragraphs.length, headings, recommendations });
  } catch (e) { return errorResponse(e); }
}