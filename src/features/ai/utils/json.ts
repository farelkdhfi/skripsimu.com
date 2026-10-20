export function extractJson(s: string) {
  const t = s
    .replace(/<think>[\s\S]*?<\/think>/g, "")
    .replace(/```json|```/g, "")
    .trim();
  const firstObj = t.indexOf("{"), firstArr = t.indexOf("[");
  const useArr = firstArr >= 0 && (firstObj < 0 || firstArr < firstObj);
  const open = useArr ? "[" : "{", close = useArr ? "]" : "}";
  const a = t.indexOf(open), b = t.lastIndexOf(close);
  return a >= 0 && b > a ? t.slice(a, b + 1) : t;
}