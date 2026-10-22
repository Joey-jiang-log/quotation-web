// Quote display rules. Keep the original catalog/model code intact for traceability.
export function normalizeRedShellInternal(internal, catalog = null) {
  const code = String(internal || "").trim().toUpperCase();
  if (!/^[A-Z][A-Z0-9-]{4,}\d{2}$/.test(code) || !catalog) return code;
  const description = [catalog.commonName, catalog.vehicle].filter(Boolean).join(" ");
  // A suffix only denotes a red-shell variant when catalog metadata explicitly says so.
  if (!/(?:红色|红壳|红外壳|red\s*(?:shell|housing|case))/i.test(description)) return code;
  return code.slice(0, -2);
}
