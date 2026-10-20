import { z } from "zod";
import { errorResponse } from "@/lib/api-response";
import { deckSchema } from "@/features/thesis-to-slides/schemas/slides";
import { buildPptx } from "@/features/thesis-to-slides/services/pptx-builder";

const body = z.object({ deck: deckSchema, style: z.enum(["academic-blue", "soft-violet", "mono-dark"]) });

export async function POST(req: Request) {
  try {
    const { deck, style } = body.parse(await req.json());
    const buf = await buildPptx(deck, style);
    return new Response(new Uint8Array(buf), {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": 'attachment; filename="presentasi-skripsi.pptx"',
      },
    });
  } catch (e) { return errorResponse(e); }
}