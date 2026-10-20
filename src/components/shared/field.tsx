import { Label } from "@/components/ui/label";

export const selectCls =
  "h-11 w-full rounded-full border border-slate-200 bg-white px-4 text-sm text-slate-900 outline-none transition focus-visible:border-purple-300 focus-visible:ring-2 focus-visible:ring-purple-200";

export const inputCls =
  "h-11 rounded-full border-slate-200 bg-white px-4 text-sm text-slate-900 shadow-none focus-visible:border-purple-300 focus-visible:ring-2 focus-visible:ring-purple-200";

export function Field({ label, error, children }: { label: string; error?: string; children: React.ReactNode }) {
  return (
    <div className="space-y-1.5">
      <Label className="pl-1 text-xs font-medium text-slate-600">{label}</Label>
      {children}
      {error && <p className="pl-1 text-xs text-red-600">{error}</p>}
    </div>
  );
}