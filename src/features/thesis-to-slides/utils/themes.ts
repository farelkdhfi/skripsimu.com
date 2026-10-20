export const THEMES = {
  "academic-blue": { name: "Academic Blue", bg: "F8FAFC", titleBg: "1E3A8A", titleText: "FFFFFF", heading: "1E3A8A", text: "1E293B", muted: "64748B", accent: "3B82F6", font: "Calibri" },
  "soft-violet": { name: "Soft Violet", bg: "FAF8FF", titleBg: "5B21B6", titleText: "FFFFFF", heading: "5B21B6", text: "2E2A3B", muted: "7C7391", accent: "8B5CF6", font: "Calibri" },
  "mono-dark": { name: "Mono Dark", bg: "111827", titleBg: "0B1220", titleText: "F9FAFB", heading: "E5E7EB", text: "D1D5DB", muted: "9CA3AF", accent: "818CF8", font: "Calibri" },
} as const;
export type ThemeKey = keyof typeof THEMES;