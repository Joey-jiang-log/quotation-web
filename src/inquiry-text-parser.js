const SKIP_WORDS = new Set([
  "PROFORMA", "INVOICE", "NUMBER", "PRICE", "TOTAL", "AMOUNT", "QUANTITY", "QTY", "PCS", "EXW", "USD", "MODEL", "PART",
]);

function normalizeCandidate(value = "") {
  return String(value).toUpperCase().replace(/^[^A-Z0-9]+|[^A-Z0-9]+$/g, "").replace(/[._/\\]+/g, "-");
}

function isLikelyModel(value = "") {
  const candidate = normalizeCandidate(value);
  const key = candidate.replace(/[^A-Z0-9]/g, "");
  if (!candidate || !/\d/.test(candidate) || key.length < 4 || key.length > 22) return false;
  if (SKIP_WORDS.has(candidate) || /^(?:19|20)\d{2}$/.test(candidate)) return false;
  if (/^\d+$/.test(key)) return key.length >= 5;
  return /[A-Z]/.test(key);
}

export function extractInquiryRows(text = "") {
  const rows = [];
  const seen = new Set();
  for (const rawLine of String(text).split(/\r?\n/)) {
    const line = rawLine.replace(/[|]+/g, " ").replace(/\s+/g, " ").trim();
    if (!line || /(?:OE\s*NUMBER|PART\s*NUMBER|QTY\s*\(?PCS|TOTAL\s*PRICE)/i.test(line)) continue;
    const matches = [...line.matchAll(/[A-Z0-9][A-Z0-9._/\\-]{3,25}/gi)];
    const modelMatch = matches.find((match) => isLikelyModel(match[0]));
    if (!modelMatch) continue;
    const model = normalizeCandidate(modelMatch[0]);
    const afterModel = line.slice((modelMatch.index || 0) + modelMatch[0].length);
    const quantity = Number(afterModel.match(/(?:^|\s)(\d{1,6})(?=\s|$)/)?.[1] || "");
    const dedupeKey = `${model}|${Number.isFinite(quantity) ? quantity : ""}|${line}`;
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);
    rows.push({ 客户型号: model, 数量: Number.isFinite(quantity) ? quantity : "", 识别原文: line });
  }
  return rows.slice(0, 500);
}
