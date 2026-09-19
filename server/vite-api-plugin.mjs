import {
  initKnowledgeBase,
  getOverview,
  searchModels,
  searchModelsBatch,
  confirmUserMapping,
  listReviews,
  reviewCandidate,
  previewQuote,
  selectQuoteCandidate,
  confirmQuoteLine,
  attachQuoteImages,
  exportQuote,
  previewKnowledgeImport,
  applyKnowledgeImport,
  applyKnowledgeFile,
  listVersions,
  syncConfirmedMappings,
  exportKnowledge,
  exportConfirmedReviews,
} from "./knowledge.mjs";

function sendJson(res, status, payload) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(payload));
}

function readBody(req, limit = 25 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > limit) {
        reject(new Error("文件不能超过 25MB"));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function fileName(req, fallback) {
  try { return decodeURIComponent(req.headers["x-file-name"] || fallback); } catch { return fallback; }
}

export function quotationApiPlugin() {
  initKnowledgeBase();
  return {
    name: "quotation-api",
    configureServer(server) {
      server.middlewares.use(async (req, res, next) => {
        if (!req.url?.startsWith("/api/")) return next();
        const url = new URL(req.url, "http://localhost");
        try {
          if (req.method === "GET" && url.pathname === "/api/overview") return sendJson(res, 200, getOverview());
          if (req.method === "GET" && url.pathname === "/api/models/search") {
            const query = url.searchParams.get("q")?.trim();
            if (!query) return sendJson(res, 400, { error: "请输入型号" });
            return sendJson(res, 200, await searchModels(query));
          }
          if (req.method === "POST" && url.pathname === "/api/models/search/batch") {
            const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
            return sendJson(res, 200, await searchModelsBatch(payload.queries));
          }
          if (req.method === "POST" && url.pathname === "/api/mappings/confirm") {
            const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
            return sendJson(res, 200, confirmUserMapping(payload.input, payload.targetInternal, payload.context, payload.reviewer));
          }
          if (req.method === "GET" && url.pathname === "/api/reviews") return sendJson(res, 200, listReviews());
          if (req.method === "GET" && url.pathname === "/api/reviews/export") {
            res.statusCode = 200;
            res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent("二次校对确认记录.xlsx")}`);
            return res.end(exportConfirmedReviews());
          }
          const reviewMatch = url.pathname.match(/^\/api\/reviews\/(\d+)$/);
          if (req.method === "POST" && reviewMatch) {
            const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
            return sendJson(res, 200, reviewCandidate(Number(reviewMatch[1]), payload.action, payload.reviewer, payload.targetInternal, payload.selectedOeNos, payload.selectedUfNos));
          }
          if (req.method === "POST" && url.pathname === "/api/quotes/preview") {
            let pricings;
            try { pricings = JSON.parse(url.searchParams.get("pricings") || "null"); } catch { pricings = null; }
            const fallbackPricing = {
              packaging: url.searchParams.get("packaging") || "彩盒",
              tax: url.searchParams.get("tax") || "未税",
              tier: url.searchParams.get("tier") || "4S",
            };
            return sendJson(res, 200, await previewQuote(await readBody(req), fileName(req, "客户询价.xlsx"), Array.isArray(pricings) ? pricings : fallbackPricing));
          }
          const quoteSelectMatch = url.pathname.match(/^\/api\/quotes\/([\w-]+)\/select$/);
          if (req.method === "POST" && quoteSelectMatch) {
            const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
            return sendJson(res, 200, selectQuoteCandidate(quoteSelectMatch[1], payload.lineIndex, payload.modelId));
          }
          const quoteConfirmMatch = url.pathname.match(/^\/api\/quotes\/([\w-]+)\/confirm$/);
          if (req.method === "POST" && quoteConfirmMatch) {
            const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
            return sendJson(res, 200, confirmQuoteLine(quoteConfirmMatch[1], payload.lineIndex, payload.targetInternal, payload.reviewer));
          }
          const quoteImagesMatch = url.pathname.match(/^\/api\/quotes\/([\w-]+)\/images$/);
          if (req.method === "POST" && quoteImagesMatch) {
            const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
            return sendJson(res, 200, attachQuoteImages(quoteImagesMatch[1], payload.images));
          }
          const quoteMatch = url.pathname.match(/^\/api\/quotes\/([\w-]+)\/export$/);
          if (req.method === "GET" && quoteMatch) {
            const buffer = await exportQuote(quoteMatch[1]);
            res.statusCode = 200;
            res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent("外贸报价结果.xlsx")}`);
            return res.end(buffer);
          }
          if (req.method === "POST" && url.pathname === "/api/knowledge/preview") {
            const requestedType = url.searchParams.get("type");
            const type = ["models", "prices", "catalog"].includes(requestedType) ? requestedType : "models";
            return sendJson(res, 200, previewKnowledgeImport(await readBody(req), fileName(req, "知识库.xlsx"), type));
          }
          if (req.method === "POST" && url.pathname === "/api/knowledge/apply") {
            const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
            return sendJson(res, 200, applyKnowledgeImport(payload.token, payload.operator));
          }
          if (req.method === "POST" && url.pathname === "/api/knowledge/import") {
            const requestedType = url.searchParams.get("type");
            const type = ["models", "prices", "catalog"].includes(requestedType) ? requestedType : "models";
            return sendJson(res, 200, applyKnowledgeFile(await readBody(req), fileName(req, "知识库.xlsx"), type, req.headers["x-operator"] || "当前用户"));
          }
          if (req.method === "GET" && url.pathname === "/api/versions") return sendJson(res, 200, listVersions());
          if (req.method === "POST" && url.pathname === "/api/knowledge/sync-confirmed") return sendJson(res, 200, syncConfirmedMappings());
          if (req.method === "GET" && url.pathname === "/api/knowledge/export") {
            const requestedType = url.searchParams.get("type");
            const type = ["models", "prices", "catalog"].includes(requestedType) ? requestedType : "models";
            const filename = type === "prices" ? "海福来-最新价格库.xlsx" : type === "catalog" ? "海福来-最新产品目录.xlsx" : "海福来-最新大厂型号匹配表.xlsx";
            res.statusCode = 200;
            res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
            return res.end(await exportKnowledge(type));
          }
          return sendJson(res, 404, { error: "接口不存在" });
        } catch (error) {
          console.error(error);
          return sendJson(res, 500, { error: error?.message || "处理失败，请稍后重试" });
        }
      });
    },
  };
}
