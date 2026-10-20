"use client";
import { useEffect, useState } from "react";
import { ChevronLeft, ChevronRight } from "lucide-react";
import { Button } from "@/components/ui/button";
import { THEMES, type ThemeKey } from "../utils/themes";
import type { Deck } from "../schemas/slides";

export function SlidePreview({ deck, style }: { deck: Deck; style: ThemeKey }) {
  const [i, setI] = useState(0);
  const t = THEMES[style];
  const total = deck.slides.length + 1;
  const go = (d: number) => setI((x) => Math.min(total - 1, Math.max(0, x + d)));

  useEffect(() => {
    const h = (e: KeyboardEvent) => { if (e.key === "ArrowRight") go(1); if (e.key === "ArrowLeft") go(-1); };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  }, []);

  const sl = i === 0 ? null : deck.slides[i - 1];
  const ch = sl?.visual.kind === "chart" ? sl.visual.chart : undefined;
  const max = ch ? Math.max(...ch.values) || 1 : 1;

  return (
    <div className="space-y-4">
      <div className="rounded-[2rem] bg-purple-100 p-3 sm:p-4">
        <div className="aspect-video overflow-hidden rounded-3xl shadow-sm" style={{ background: `#${i === 0 ? t.titleBg : t.bg}`, color: `#${t.text}` }}>
          {i === 0 ? (
            <div className="flex h-full flex-col justify-center p-[6%]" style={{ color: `#${t.titleText}` }}>
              <h2 className="text-xl font-bold leading-tight sm:text-3xl">{deck.title}</h2>
              <div className="my-3 h-1 w-12 rounded" style={{ background: `#${t.accent}` }} />
              <p className="text-xs opacity-90 sm:text-base">{deck.subtitle}</p>
            </div>
          ) : (
            <div className="flex h-full gap-3 p-[4%] pl-[6%]">
              <div className="flex-1 overflow-hidden">
                <h3 className="mb-3 text-base font-bold sm:text-2xl" style={{ color: `#${t.heading}` }}>{sl!.title}</h3>
                <ul className="list-disc space-y-1.5 pl-5 text-[11px] sm:text-base">{sl!.bullets.map((b, k) => <li key={k}>{b}</li>)}</ul>
              </div>
              {ch && (
                <div className="w-2/5 space-y-1.5 self-center text-[10px] sm:text-xs">
                  {ch.title && <p className="font-semibold">{ch.title}</p>}
                  {ch.labels.map((l, k) => (
                    <div key={k}>
                      <div className="flex justify-between"><span>{l}</span><span>{ch.values[k]}</span></div>
                      <div className="h-2 rounded" style={{ width: `${(ch.values[k] / max) * 100}%`, background: `#${t.accent}` }} />
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      <div className="flex items-center justify-between">
        <Button
          variant="outline"
          size="icon"
          onClick={() => go(-1)}
          disabled={i === 0}
          aria-label="Sebelumnya"
          className="h-10 w-10 rounded-full border-slate-200 bg-white text-slate-900 hover:bg-slate-50"
        >
          <ChevronLeft className="size-4" />
        </Button>
        <span className="rounded-full bg-purple-100 px-3 py-1 text-xs font-medium text-purple-700">Slide {i + 1} / {total}</span>
        <Button
          variant="outline"
          size="icon"
          onClick={() => go(1)}
          disabled={i === total - 1}
          aria-label="Berikutnya"
          className="h-10 w-10 rounded-full border-slate-200 bg-white text-slate-900 hover:bg-slate-50"
        >
          <ChevronRight className="size-4" />
        </Button>
      </div>

      {sl && (sl.notes || sl.visual.suggestion) && (
        <div className="space-y-1.5 rounded-[1.5rem] bg-green-100 p-5 text-sm leading-relaxed text-slate-700">
          {sl.notes && <p><span className="font-medium text-slate-900">Catatan presenter:</span> {sl.notes}</p>}
          {sl.visual.suggestion && <p><span className="font-medium text-slate-900">Saran visual:</span> {sl.visual.suggestion}</p>}
        </div>
      )}
    </div>
  );
}