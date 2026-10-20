import { NextResponse } from "next/server";
import { errorResponse } from "@/lib/api-response";
import { parseUpload } from "@/features/file-upload/services/parse-upload";
import { slidesConfigSchema } from "@/features/thesis-to-slides/schemas/slides";
import { generateDeck } from "@/features/thesis-to-slides/services/slides.service";

export const maxDuration = 300;

export async function POST(req: Request) {
  try {
    const { file, fd } = await parseUpload(req);
    const config = slidesConfigSchema.parse(JSON.parse(String(fd.get("config") ?? "{}")));
    return NextResponse.json(await generateDeck(file, config));
  } catch (e) { return errorResponse(e); }
}