// Quote display rules. Keep the original catalog/model code intact for traceability.
export function normalizeRedShellInternal(internal, catalog = null) {
  const source = String(internal || "").trim();
  const code = source.toUpperCase();
  if (!/^[A-Z][A-Z0-9-]{4,}\d{2}$/.test(code) || !catalog || String(catalog.internal || "").trim().toUpperCase() !== code) return source;
  const description = [catalog.commonName, catalog.vehicle].filter(Boolean).join(" ");
  // A suffix only denotes a red-shell variant when catalog metadata explicitly says so.
  if (!/(?:红色|红壳|红外壳|red\s*(?:shell|housing|case))/i.test(description)) return source;
  return source.slice(0, -2);
}
