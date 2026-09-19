const SEARCH_TIMEOUT_MS = 4500;

const CROSS_REFERENCE_CATALOGS = [
  {
    platform: "Alibaba",
    url: "https://www.alibaba.com/ignition-coil-automotive-parts-manufacturers.html",
  },
  {
    platform: "Alibaba",
    url: "https://www.alibaba.com/ignition-coil-importers-suppliers.html",
  },
  {
    platform: "MachineryOffers",
    url: "https://www.machineryoffers.com/sp/vx-ignition-coils",
  },
];

const catalogDocumentCache = new Map();

function decodeHtml(value = "") {
  return value
    .replace(/<[^>]+>/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#x27;|&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/\s+/g, " ")
    .trim();
}

function unwrapDuckDuckGoUrl(value = "") {
  const decoded = decodeHtml(value);
  try {
    const url = new URL(decoded, "https://duckduckgo.com");
    const target = url.searchParams.get("uddg");
    return target ? decodeURIComponent(target) : url.href;
  } catch {
    return decoded;
  }
}

function unique(values) {
  return [...new Set(values.filter(Boolean))];
}

function normalizeCandidate(value = "") {
  return String(value).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function isUsefulExternalCandidate(value = "") {
  const text = String(value).trim().toUpperCase();
  const key = normalizeCandidate(text);
  if (!text || /^DQ\d/.test(text)) return false;
  if (/^(?:19|20)\d{2}(?:\s*[-–—/]\s*(?:19|20)\d{2})?$/.test(text)) return false;
  if (/^\d{1,4}(?:\.\d+)?$/.test(text)) return false;
  return /\d/.test(text) && key.length >= 5;
}

function compactEvidenceText(value = "", maxLength = 520) {
  const clean = decodeHtml(String(value)
    .replace(/!\[[^\]]*\]\([^)]*\)/g, " ")
    .replace(/\[([^\]]+)\]\(https?:\/\/[^)]+\)/g, "$1")
    .replace(/(?:Markdown Content|URL Source|Published Time|Title):/gi, " "));
  return clean.length > maxLength ? `${clean.slice(0, maxLength).trim()}…` : clean;
}

function extractCandidates(text = "") {
  const cleanText = text.replace(/https?:\/\/\S+/gi, " ").replace(/%[0-9A-F]{2}/gi, " ");
  const ufNos = unique((cleanText.match(/\bUF[-\s]?(?=[A-Z0-9]{2,8}\b)(?=[A-Z0-9]*\d)[A-Z0-9]{2,8}\b/gi) || []).map((value) => value.replace(/\s+/g, "").toUpperCase()));
  const oeNos = unique(
    (cleanText.match(/\b(?=[A-Z0-9-]{6,20}\b)(?=[A-Z0-9-]*\d{4})[A-Z0-9]+(?:-[A-Z0-9]+)*\b/gi) || [])
      .map((value) => value.toUpperCase())
      .filter((value) => !/^UF/.test(value) && !/^(IGNITION|COIL|AMAZON|EBAY)$/i.test(value) && isUsefulExternalCandidate(value)),
  ).slice(0, 12);
  return { ufNos, oeNos };
}

async function fetchWithTimeout(url, options = {}, timeoutMs = SEARCH_TIMEOUT_MS) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, {
      ...options,
      signal: controller.signal,
      headers: {
        "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/126 Safari/537.36",
        "accept-language": "zh-CN,zh;q=0.9,en;q=0.8",
        ...(options.headers || {}),
      },
    });
  } finally {
    clearTimeout(timer);
  }
}

async function braveSearch(searchText) {
  const apiKey = process.env.BRAVE_SEARCH_API_KEY;
  if (!apiKey) return [];
  const response = await fetchWithTimeout(`https://api.search.brave.com/res/v1/web/search?q=${encodeURIComponent(searchText)}&count=6`, {
    headers: { Accept: "application/json", "X-Subscription-Token": apiKey },
  });
  if (!response.ok) throw new Error(`Brave Search 返回 ${response.status}`);
  const data = await response.json();
  return (data.web?.results || []).map((item) => ({
    title: item.title || "外网搜索结果",
    url: item.url,
    snippet: decodeHtml(item.description || ""),
    provider: "Brave Search",
  }));
}

async function duckDuckGoSearch(searchText) {
  const response = await fetchWithTimeout(`https://html.duckduckgo.com/html/?q=${encodeURIComponent(searchText)}`);
  if (!response.ok) throw new Error(`DuckDuckGo 返回 ${response.status}`);
  const html = await response.text();
  const resultRegex = /<a[^>]*class="[^"]*result__a[^"]*"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>[\s\S]*?(?:class="[^"]*result__snippet[^"]*"[^>]*>([\s\S]*?)<\/a>|class="[^"]*result__snippet[^"]*"[^>]*>([\s\S]*?)<\/div>)/gi;
  const results = [];
  let match;
  while ((match = resultRegex.exec(html)) && results.length < 6) {
    results.push({
      title: decodeHtml(match[2]),
      url: unwrapDuckDuckGoUrl(match[1]),
      snippet: decodeHtml(match[3] || match[4] || ""),
      provider: "DuckDuckGo",
    });
  }
  if (!results.length) {
    const anchorRegex = /<a[^>]*class="[^"]*result__a[^"]*"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi;
    while ((match = anchorRegex.exec(html)) && results.length < 6) {
      results.push({
        title: decodeHtml(match[2]),
        url: unwrapDuckDuckGoUrl(match[1]),
        snippet: "",
        provider: "DuckDuckGo",
      });
    }
  }
  return results;
}

async function googleSearch(searchText) {
  const target = `http://www.google.com/search?hl=en&num=8&q=${encodeURIComponent(searchText)}`;
  const reader = await fetchWithTimeout(`https://r.jina.ai/${target}`, {}, 8500);
  if (reader.ok) {
    const text = await reader.text();
    const blocked = /Target URL returned error|Error Page|unusual traffic|CAPTCHA|not a robot|automated queries/i.test(text);
    if (text.trim() && !blocked) return [{ title: `Google：${searchText}`, url: target.replace("http://", "https://"), snippet: text.slice(0, 30000), provider: "Google", platform: "Google" }];
  }
  const response = await fetchWithTimeout(target.replace("http://", "https://"));
  if (!response.ok) throw new Error(`Google 返回 ${response.status}`);
  const html = await response.text();
  const results = [];
  const regex = /<a[^>]+href="\/url\?q=([^"&]+)[^"]*"[^>]*>[\s\S]*?<h3[^>]*>([\s\S]*?)<\/h3>/gi;
  let match;
  while ((match = regex.exec(html)) && results.length < 6) {
    const url = decodeURIComponent(match[1]);
    if (!/^https?:/i.test(url)) continue;
    results.push({ title: decodeHtml(match[2]), url, snippet: "", provider: "Google", platform: "Google" });
  }
  if (results.length) return results;
  // Google 直连在部分云环境不返回正文时，使用公开网页索引补足结果条目，仍按 Google 校对通道展示。
  const fallback = await duckDuckGoSearch(searchText);
  return fallback.map((item) => ({ ...item, provider: "Google fallback index", platform: "Google" }));
}

async function amazonSearch(searchText) {
  const target = `http://www.amazon.com/s?k=${encodeURIComponent(searchText)}`;
  try {
    const reader = await fetchWithTimeout(`https://r.jina.ai/${target}`, {}, 8500);
    if (reader.ok) {
      const text = await reader.text();
      if (text.trim() && !/Target URL returned error|Error Page/i.test(text)) return [{ title: `Amazon：${searchText}`, url: target.replace("http://", "https://"), snippet: text.slice(0, 30000), provider: "Amazon", platform: "Amazon" }];
    }
  } catch { /* 继续使用站内索引搜索 */ }
  const fallback = await duckDuckGoSearch(`site:amazon.com ${searchText}`);
  return fallback.map((item) => ({ ...item, provider: "Amazon via Web Search", platform: "Amazon" }));
}

async function ebaySearch(searchText) {
  const target = `http://www.ebay.com/sch/i.html?_nkw=${encodeURIComponent(searchText)}`;
  try {
    const reader = await fetchWithTimeout(`https://r.jina.ai/${target}`, {}, 8500);
    if (reader.ok) {
      const text = await reader.text();
      if (text.trim() && !/Target URL returned error|Error Page|SOMETHING WENT WRONG/i.test(text)) return [{ title: `eBay：${searchText}`, url: target.replace("http://", "https://"), snippet: text.slice(0, 30000), provider: "eBay", platform: "eBay" }];
    }
  } catch { /* 继续使用站内索引搜索 */ }
  const fallback = await duckDuckGoSearch(`site:ebay.com ${searchText}`);
  return fallback.map((item) => ({ ...item, provider: "eBay via Web Search", platform: "eBay" }));
}

async function getCatalogDocument(catalog) {
  const cached = catalogDocumentCache.get(catalog.url);
  if (cached) return cached;
  const pending = fetchWithTimeout(catalog.url, {}, 12000)
    .then(async (response) => {
      if (!response.ok) throw new Error(`${catalog.platform} 返回 ${response.status}`);
      return response.text();
    })
    .catch(() => "");
  catalogDocumentCache.set(catalog.url, pending);
  return pending;
}

function findEvidenceWindows(document, needles) {
  const upperDocument = document.toUpperCase();
  const windows = [];
  for (const needle of needles) {
    const normalizedNeedle = String(needle || "").trim().toUpperCase();
    if (!normalizedNeedle) continue;
    let offset = 0;
    while (windows.length < 6) {
      const index = upperDocument.indexOf(normalizedNeedle, offset);
      if (index < 0) break;
      windows.push(document.slice(Math.max(0, index - 260), Math.min(document.length, index + normalizedNeedle.length + 360)));
      offset = index + normalizedNeedle.length;
    }
  }
  return unique(windows.map(decodeHtml)).join(" · ");
}

async function catalogCrossReferenceSearch(query, context = {}) {
  const needles = unique([query, ...(context.oeNos || []).slice(0, 4), ...(context.ufNos || []).slice(0, 2)]);
  const matches = await Promise.all(CROSS_REFERENCE_CATALOGS.map(async (catalog) => {
    const document = await getCatalogDocument(catalog);
    const snippet = document ? findEvidenceWindows(document, needles) : "";
    if (!snippet) return null;
    return {
      title: `${catalog.platform} 配件目录交叉参考`,
      url: catalog.url,
      snippet: snippet.slice(0, 2000),
      provider: `${catalog.platform} Catalog`,
      platform: catalog.platform,
    };
  }));
  return matches.filter(Boolean);
}

async function searchPlatform(platform, searchText) {
  const searcher = platform === "Google" ? googleSearch : platform === "Amazon" ? amazonSearch : ebaySearch;
  try {
    const results = await searcher(searchText);
    return { platform, results: results.map((item) => ({ ...item, platform })), error: results.length ? "" : "未找到有效结果" };
  } catch (error) {
    return { platform, results: [], error: error?.name === "AbortError" ? "搜索超时" : error?.message || "搜索失败" };
  }
}

export async function searchExternalModels(query, context = {}) {
  const knownOe = (context.oeNos || []).slice(0, 3).join(" ");
  const primaryOe = (context.oeNos || [])[0] || query;
  const searchText = `"${query}" ${knownOe} ignition coil UF OE cross reference`.replace(/\s+/g, " ").trim();
  const marketplaceText = `"${primaryOe}" ignition coil UF`.trim();
  const ufSearchTexts = unique([context.internal, query, primaryOe])
    .slice(0, 3)
    .map((anchor) => `"${anchor}" ignition coil "UF" cross reference`);
  const [platformResults, catalogResults] = await Promise.all([
    Promise.all([
      searchPlatform("Google", searchText),
      ...ufSearchTexts.map((text) => searchPlatform("Google", text)),
      searchPlatform("Amazon", marketplaceText),
      searchPlatform("eBay", marketplaceText),
    ]),
    catalogCrossReferenceSearch(query, context),
  ]);
  let results = [...platformResults.flatMap((item) => item.results), ...catalogResults];
  if (!results.length) {
    try {
      const fallback = await braveSearch(searchText);
      results = fallback.map((item) => ({ ...item, platform: "Web" }));
    } catch {
      results = [];
    }
  }
  const deduped = [];
  const seen = new Set();
  for (const item of results) {
    if (!item.url || seen.has(item.url)) continue;
    seen.add(item.url);
    deduped.push(item);
  }
  const support = { oe: new Map(), uf: new Map() };
  const evidence = [];
  for (const item of deduped) {
    const fullEvidence = `${item.title} ${item.snippet}`;
    const anchors = unique([query, ...(context.oeNos || []), context.internal]);
    const anchoredEvidence = findEvidenceWindows(fullEvidence, anchors) || item.title;
    const candidates = extractCandidates(anchoredEvidence);
    const oeNos = candidates.oeNos.filter((value) => normalizeCandidate(value) !== normalizeCandidate(query));
    const ufNos = candidates.ufNos.filter((value) => normalizeCandidate(value) !== normalizeCandidate(query));
    evidence.push({ platform: item.platform || item.provider, url: item.url, oeNos, ufNos });
    for (const [kind, values] of [["oe", oeNos], ["uf", ufNos]]) {
      for (const value of values) {
        const key = normalizeCandidate(value);
        const record = support[kind].get(key) || { value, platforms: new Set() };
        record.platforms.add(item.platform || item.provider || "Web");
        support[kind].set(key, record);
      }
    }
  }
  const flatten = (map) => [...map.values()].sort((a, b) => {
    if (b.platforms.size !== a.platforms.size) return b.platforms.size - a.platforms.size;
    return Number(b.platforms.has("Google")) - Number(a.platforms.has("Google"));
  }).map((item) => ({ value: item.value, platforms: [...item.platforms], verified: item.platforms.size >= 2 }));
  // 单独出现在 Amazon / eBay 搜索页的 UF 往往来自推荐商品，不能视为当前 OE 的候选。
  const keepCandidate = (item) => item.verified || item.platforms.includes("Google");
  const oeEvidence = flatten(support.oe).filter(keepCandidate);
  const ufEvidence = flatten(support.uf).filter(keepCandidate).slice(0, 2);
  const activePlatforms = unique(results.map((item) => item.platform || item.provider));
  const verifiedCount = [...oeEvidence, ...ufEvidence].filter((item) => item.verified).length;
  const confidence = verifiedCount ? 88 : activePlatforms.length >= 2 && (oeEvidence.length || ufEvidence.length) ? 68 : activePlatforms.length ? 48 : 0;
  return {
    query,
    provider: activePlatforms.join(" + ") || "Web Search",
    searchedAt: new Date().toISOString(),
    sources: deduped.slice(0, 12).map((item) => ({ ...item, snippet: compactEvidenceText(item.snippet || item.url) })),
    oeNos: oeEvidence.filter((item) => item.verified).map((item) => item.value),
    // 最多保留两个有 Google 或跨平台证据的 UF，仍需用户逐个确认。
    ufNos: ufEvidence.map((item) => item.value),
    verifiedOeNos: oeEvidence.filter((item) => item.verified).map((item) => item.value),
    verifiedUfNos: ufEvidence.filter((item) => item.verified).map((item) => item.value),
    candidateEvidence: { oe: oeEvidence, uf: ufEvidence },
    platforms: activePlatforms,
    platformErrors: Object.fromEntries(platformResults.filter((item) => item.error).map((item) => [item.platform, item.error])),
    evidence,
    confidence,
    error: deduped.length ? "" : "Google、Amazon、eBay 及公开配件目录均未找到可引用结果",
  };
}
