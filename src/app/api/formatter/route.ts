import { NextResponse } from "next/server";
import { errorResponse } from "@/lib/api-response";
import { parseUpload } from "@/features/file-upload/services/parse-upload";
import { parseFormatConfig } from "@/features/thesis-formatter/schemas/format-config";
import { formatThesis } from "@/features/thesis-formatter/services/formatter.service";

export const maxDuration = 180;

export async function POST(req: Request) {
  try {
    const { file, fd } = await parseUpload(req);
    const config = parseFormatConfig(fd.get("config"));
    return NextResponse.json(await formatThesis(file, config));
  } catch (e) { return errorResponse(e); }
}