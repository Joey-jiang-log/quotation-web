async function parseResponse(response) {
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || "请求失败，请稍后重试");
  return data;
}

export const api = {
  overview: () => fetch("/api/overview", { cache: "no-store" }).then(parseResponse),
  searchModels: (query) => fetch(`/api/models/search?q=${encodeURIComponent(query)}`).then(parseResponse),
  searchModelsBatch: (queries) => fetch("/api/models/search/batch", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ queries }),
  }).then(parseResponse),
  confirmMapping: (input, targetInternal, context) => fetch("/api/mappings/confirm", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ input, targetInternal, context, reviewer: "当前用户" }),
  }).then(parseResponse),
  reviews: () => fetch("/api/reviews", { cache: "no-store" }).then(parseResponse),
  review: (id, action, reviewer = "当前用户", targetInternal = "", selectedOeNos = [], selectedUfNos = []) => fetch(`/api/reviews/${id}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ action, reviewer, targetInternal, selectedOeNos, selectedUfNos }),
  }).then(parseResponse),
  previewQuote: (file, pricings) => fetch(`/api/quotes/preview?${new URLSearchParams({ pricings: JSON.stringify(pricings) })}`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream", "X-File-Name": encodeURIComponent(file.name) },
    body: file,
  }).then(parseResponse),
  selectQuoteCandidate: (jobId, lineIndex, modelId) => fetch(`/api/quotes/${jobId}/select`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ lineIndex, modelId }),
  }).then(parseResponse),
  confirmQuoteLine: (jobId, lineIndex, targetInternal) => fetch(`/api/quotes/${jobId}/confirm`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ lineIndex, targetInternal, reviewer: "当前用户" }),
  }).then(parseResponse),
  attachQuoteImages: (jobId, images) => fetch(`/api/quotes/${jobId}/images`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ images }),
  }).then(parseResponse),
  previewKnowledge: (file, type) => fetch(`/api/knowledge/preview?type=${type}`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream", "X-File-Name": encodeURIComponent(file.name) },
    body: file,
  }).then(parseResponse),
  applyKnowledge: (token) => fetch("/api/knowledge/apply", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ token, operator: "当前用户" }),
  }).then(parseResponse),
  importKnowledge: (file, type) => fetch(`/api/knowledge/import?type=${type}`, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream", "X-File-Name": encodeURIComponent(file.name), "X-Operator": "当前用户" },
    body: file,
  }).then(parseResponse),
  versions: () => fetch("/api/versions").then(parseResponse),
};
