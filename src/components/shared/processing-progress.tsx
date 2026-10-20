"use client";
import { useEffect, useState } from "react";
import { Loader2 } from "lucide-react";
import { Card, CardContent } from "@/components/ui/card";
import { Progress } from "@/components/ui/progress";

export function ProcessingProgress({ active, stages }: { active: boolean; stages: string[] }) {
  const [v, setV] = useState(0);
  useEffect(() => {
    if (!active) { setV(0); return; }
    const id = setInterval(() => setV((x) => x + (92 - x) * 0.04), 500);
    return () => clearInterval(id);
  }, [active]);
  if (!active) return null;
  const stage = stages[Math.min(stages.length - 1, Math.floor(v / (92 / stages.length)))];
  return (
    <Card className="rounded-[2rem] border-0 bg-slate-900 text-white shadow-none">
      <CardContent className="space-y-4 p-6">
        <div className="flex items-center gap-3 text-sm">
          <span className="grid size-8 place-items-center rounded-full bg-white/10">
            <Loader2 className="size-4 animate-spin" />
          </span>
          {stage}
        </div>
        <Progress value={v} className="h-2 bg-white/15 [&>*]:bg-purple-300" />
        <p className="text-xs text-slate-400">Perkiraan progres. Dokumen panjang diproses bertahap dan bisa memakan waktu beberapa menit.</p>
      </CardContent>
    </Card>
  );
}