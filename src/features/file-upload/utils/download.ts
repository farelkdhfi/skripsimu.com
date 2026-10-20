export function downloadBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement("a"), { href: url, download: name });
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
export function base64ToBlob(b64: string, type: string) {
  const bin = atob(b64);
  const arr = Uint8Array.from(bin, (c) => c.charCodeAt(0));
  return new Blob([arr], { type });
}