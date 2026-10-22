import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { randomUUID } from "node:crypto";
import { gzipSync, gunzipSync } from "node:zlib";
import { DatabaseSync } from "node:sqlite";
import { neon } from "@neondatabase/serverless";
import XLSX from "xlsx";
import ExcelJS from "exceljs";
import { searchExternalModels } from "./search.mjs";
import { normalizeRedShellInternal } from "./red-shell.mjs";

const SEARCH_PROVIDER_PREFIX = "Google UF 专项校对 v5";
const MAX_BATCH_MODEL_QUERIES = 500;
const MAX_WEB_BATCH_QUERIES = 40;

const rootDir = process.cwd();
const sourceDataDir = path.join(rootDir, "data");
const dataDir = process.env.VERCEL ? path.join(os.tmpdir(), "hifly-quotation-data") : sourceDataDir;
const uploadDir = path.join(dataDir, "uploads");
const dbPath = path.join(dataDir, "quotation.sqlite");
const sourceModelPath = path.join(sourceDataDir, "source", "大厂型号匹配260809.xlsx");
const sourcePricePath = path.join(sourceDataDir, "source", "外贸价格体系60319.xlsx");
const sourceCatalogPath = path.join(sourceDataDir, "source", "产品目录60625.json");
const publicDir = path.join(rootDir, "public");

fs.mkdirSync(uploadDir, { recursive: true });
const db = new DatabaseSync(dbPath);
db.exec("PRAGMA journal_mode=WAL; PRAGMA foreign_keys=ON;");

function normalize(value = "") {
  return String(value).trim().toUpperCase().replace(/[\s._/\\-]+/g, "").replace(/[^A-Z0-9]/g, "");
}

function json(value, fallback = []) {
  try { return JSON.parse(value); } catch { return fallback; }
}

function unique(values) {
  return [...new Set(values.map((value) => String(value || "").trim()).filter(Boolean))];
}

function mergeTextValues(...values) {
  return unique(values.flatMap((value) => String(value || "").split(/\s*(?:\/|,|，|;|；|\r?\n)\s*/))).join(" / ");
}

function compactReviewSource(source = {}) {
  const clean = String(source.snippet || source.url || "")
    .replace(/!\[[^\]]*\]\([^)]*\)/g, " ")
    .replace(/\[([^\]]+)\]\(https?:\/\/[^)]+\)/g, "$1")
    .replace(/<[^>]+>/g, " ")
    .replace(/(?:Markdown Content|URL Source|Published Time|Title):/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
  return { ...source, snippet: clean.length > 520 ? `${clean.slice(0, 520)}…` : clean };
}

function productImagesFromSources(sources = []) {
  return [...new Map(sources.flatMap((source) => (source.imageUrls || []).map((url) => [url, {
    url,
    platform: source.platform || source.provider || "Web",
    sourceUrl: source.url,
    title: source.title || "外网产品图片",
  }])).values())].slice(0, 9);
}

function parseIdentifierList(value) {
  if (!value) return [];
  const tokens = String(value).toUpperCase().match(/[A-Z0-9]+(?:-[A-Z0-9]+)*/g) || [];
  return unique(tokens)
    .filter((item) => {
      const key = normalize(item);
      const hasLetter = /[A-Z]/.test(item);
      const hasDigit = /\d/.test(item);
      if (!hasDigit || /^\d{4}-\d{4}$/.test(item) || /^(?:19|20)\d{2}$/.test(item)) return false;
      return hasLetter ? key.length >= 4 : key.length >= 5;
    })
    .slice(0, 80);
}

function isValidModelInput(value) {
  const text = String(value || "").trim().toUpperCase();
  const key = normalize(text);
  if (!text || /^(?:QTY|QUANTITY|PCS|数量|单价|价格|TOTAL|AMOUNT|NO|序号)$/i.test(text)) return false;
  if (/^\d{1,4}(?:\.\d+)?$/.test(text) || /^\d{4}\s*[-–—]\s*\d{4}$/.test(text) || /^(?:19|20)\d{2}$/.test(text)) return false;
  if (!/\d/.test(text)) return false;
  return /[A-Z]/.test(text) ? key.length >= 4 : key.length >= 5;
}

function splitUniversalInputs(value) {
  const source = String(value || "").trim();
  if (!source) return [];
  const parts = source
    .replace(/[()（）]/g, " ")
    .split(/\s*(?:\/|,|，|;|；|\||\r?\n|\bor\b|\band\b)\s*/i)
    .map((item) => item.trim().replace(/^[,;，；|]+|[,;，；|]+$/g, ""))
    .map((item) => item.replace(/^[A-Z][A-Z0-9 _-]{1,24}[:：]\s*(?=[A-Z0-9]*\d)/i, "").trim())
    .filter(isValidModelInput);
  return unique(parts.length ? parts : [source]);
}

function expandUniversalValues(values = []) {
  return unique(values.flatMap((value) => {
    const parts = splitUniversalInputs(value);
    return parts.length > 1 ? parts : [value];
  }));
}

function readWorkbook(bufferOrPath) {
  return typeof bufferOrPath === "string"
    ? XLSX.readFile(bufferOrPath, { cellFormula: true, cellText: false, cellDates: true })
    : XLSX.read(bufferOrPath, { type: "buffer", cellFormula: true, cellText: false, cellDates: true });
}

function parseModels(bufferOrPath) {
  const workbook = readWorkbook(bufferOrPath);
  const collected = new Map();
  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: false });
    const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell).includes("对外型号")));
    if (headerIndex < 0) continue;
    const headers = rows[headerIndex].map((cell) => String(cell).trim());
    const indexes = {
      product: headers.findIndex((header) => header.includes("货品编码")),
      internal: headers.findIndex((header) => header.includes("对外型号")),
      old: headers.findIndex((header) => header.includes("老型号")),
      oe: headers.findIndex((header) => /OE号/i.test(header)),
      uf: headers.findIndex((header) => /UF号/i.test(header)),
      vehicle: headers.findIndex((header) => header.includes("适用车型")),
      category: headers.findIndex((header) => header.includes("货品类别")),
    };
    for (const row of rows.slice(headerIndex + 1)) {
      const internal = indexes.internal >= 0 ? String(row[indexes.internal] || "").trim() : "";
      if (!internal) continue;
      const key = normalize(internal);
      const current = collected.get(key) || {
        internal,
        productCodes: [],
        oldModels: [],
        oeNos: [],
        ufNos: [],
        vehicle: "",
        category: "",
        sourceSheet: sheetName,
      };
      current.productCodes = unique([...current.productCodes, indexes.product >= 0 ? row[indexes.product] : ""]);
      current.oldModels = unique([...current.oldModels, indexes.old >= 0 ? row[indexes.old] : ""]);
      const parsedOe = parseIdentifierList(indexes.oe >= 0 ? row[indexes.oe] : "");
      const parsedUf = parseIdentifierList(indexes.uf >= 0 ? row[indexes.uf] : "");
      current.oeNos = unique([...current.oeNos, ...parsedOe.filter((item) => !/^UF/i.test(item))]);
      current.ufNos = unique([...current.ufNos, ...parsedUf, ...parsedOe.filter((item) => /^UF/i.test(item))]);
      current.vehicle = current.vehicle || (indexes.vehicle >= 0 ? String(row[indexes.vehicle] || "").trim() : "");
      current.category = current.category || (indexes.category >= 0 ? String(row[indexes.category] || "").trim() : "");
      current.sourceSheet = sheetName;
      collected.set(key, current);
    }
  }
  return [...collected.values()];
}

const PRICE_SHEET_MAP = {
  "彩含价": { packaging: "彩盒", tax: "含税" },
  "彩未价": { packaging: "彩盒", tax: "未税" },
  "中含价": { packaging: "中性", tax: "含税" },
  "中未价": { packaging: "中性", tax: "未税" },
};

function parsePrices(bufferOrPath) {
  const workbook = readWorkbook(bufferOrPath);
  const prices = [];
  for (const sheetName of workbook.SheetNames) {
    const mappingKey = Object.keys(PRICE_SHEET_MAP).find((key) => sheetName.startsWith(key));
    if (!mappingKey) continue;
    const context = PRICE_SHEET_MAP[mappingKey];
    const sheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    const headers = [];
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      headers[col] = String(sheet[XLSX.utils.encode_cell({ r: 0, c: col })]?.v || "").trim();
    }
    const internalCol = headers.findIndex((header) => header.includes("对外型号"));
    const attributeCol = headers.findIndex((header) => header.includes("属性"));
    const tierCols = {};
    for (const tier of ["NP", "LP", "SLP", "2S", "3S", "4S"]) {
      tierCols[tier] = headers.findIndex((header) => header.toUpperCase().startsWith(tier));
    }
    for (let row = 1; row <= range.e.r; row += 1) {
      const internal = String(sheet[XLSX.utils.encode_cell({ r: row, c: internalCol })]?.v || "").trim();
      if (!internal) continue;
      const attribute = attributeCol >= 0 ? String(sheet[XLSX.utils.encode_cell({ r: row, c: attributeCol })]?.v || "").trim() : "";
      for (const [tier, col] of Object.entries(tierCols)) {
        if (col < 0) continue;
        const numeric = Number(sheet[XLSX.utils.encode_cell({ r: row, c: col })]?.v);
        if (!Number.isFinite(numeric)) continue;
        prices.push({ internal, ...context, tier, price: numeric, attribute, sourceSheet: sheetName, sourceColumn: headers[col] });
      }
    }
  }
  return prices;
}

function reviewModelCodes(value) {
  return unique([...String(value || "").matchAll(/\bDQ[A-Z0-9-]*\d[A-Z0-9-]*\b/gi)].map((match) => match[0].toUpperCase()));
}

function parseMappingReview(bufferOrPath) {
  const workbook = readWorkbook(bufferOrPath);
  const decisions = [];
  const mergedRelations = [];
  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: false });
    const explicitHeaderIndex = rows.findIndex((row) => {
      const headers = row.map((cell) => String(cell).trim());
      return headers.some((header) => header.includes("确认型号")) && headers.some((header) => header.includes("备注")) && headers.some((header) => /OE号/i.test(header));
    });
    if (explicitHeaderIndex >= 0) {
      const headers = rows[explicitHeaderIndex].map((cell) => String(cell).trim());
      const oeCol = headers.findIndex((header) => /OE号/i.test(header));
      const confirmCol = headers.findIndex((header) => header.includes("确认型号"));
      const noteCol = headers.findIndex((header) => header.includes("备注"));
      for (const row of rows.slice(explicitHeaderIndex + 1)) {
        const oeNos = parseIdentifierList(row[oeCol]);
        if (!oeNos.length) continue;
        const confirmModels = reviewModelCodes(row[confirmCol]);
        const note = String(row[noteCol] || "").trim();
        decisions.push({ oeNos, confirmModels, note, sheetName, source: "备注核验" });
      }
      continue;
    }
    const mergedHeaderIndex = rows.findIndex((row) => {
      const headers = row.map((cell) => String(cell).trim());
      return headers.some((header) => header.includes("合并后内部型号")) && headers.some((header) => /OE号/i.test(header));
    });
    if (mergedHeaderIndex < 0) continue;
    const headers = rows[mergedHeaderIndex].map((cell) => String(cell).trim());
    const oeCol = headers.findIndex((header) => /OE号/i.test(header));
    const mergedCol = headers.findIndex((header) => header.includes("合并后内部型号"));
    for (const row of rows.slice(mergedHeaderIndex + 1)) {
      const oeNos = parseIdentifierList(row[oeCol]);
      const modelCodes = reviewModelCodes(row[mergedCol]);
      if (oeNos.length && modelCodes.length) mergedRelations.push({ oeNos, modelCodes, sheetName, source: "双方合并复核" });
    }
  }
  return { decisions, mergedRelations };
}

function setupSchema() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS models (
      id INTEGER PRIMARY KEY AUTOINCREMENT, internal TEXT NOT NULL, normalized_internal TEXT NOT NULL UNIQUE,
      product_codes TEXT NOT NULL DEFAULT '[]', old_models TEXT NOT NULL DEFAULT '[]', oe_nos TEXT NOT NULL DEFAULT '[]', uf_nos TEXT NOT NULL DEFAULT '[]',
      web_oe_nos TEXT NOT NULL DEFAULT '[]', web_uf_nos TEXT NOT NULL DEFAULT '[]',
      vehicle TEXT NOT NULL DEFAULT '', category TEXT NOT NULL DEFAULT '', source_type TEXT NOT NULL, source_label TEXT NOT NULL,
      source_url TEXT NOT NULL DEFAULT '', captured_at TEXT NOT NULL, confidence INTEGER NOT NULL DEFAULT 100, review_status TEXT NOT NULL DEFAULT '已确认'
    );
    CREATE TABLE IF NOT EXISTS identifiers (
      id INTEGER PRIMARY KEY AUTOINCREMENT, model_id INTEGER NOT NULL REFERENCES models(id) ON DELETE CASCADE,
      kind TEXT NOT NULL, raw_value TEXT NOT NULL, normalized_value TEXT NOT NULL, source_type TEXT NOT NULL DEFAULT 'original', UNIQUE(model_id, kind, normalized_value)
    );
    CREATE INDEX IF NOT EXISTS idx_identifiers_value ON identifiers(normalized_value);
    CREATE TABLE IF NOT EXISTS prices (
      id INTEGER PRIMARY KEY AUTOINCREMENT, normalized_internal TEXT NOT NULL, packaging TEXT NOT NULL, tax TEXT NOT NULL, tier TEXT NOT NULL,
      price REAL NOT NULL, attribute TEXT NOT NULL DEFAULT '', source_sheet TEXT NOT NULL, source_column TEXT NOT NULL, UNIQUE(normalized_internal, packaging, tax, tier)
    );
    CREATE TABLE IF NOT EXISTS confirmed_relations (
      id INTEGER PRIMARY KEY AUTOINCREMENT, model_id INTEGER NOT NULL REFERENCES models(id) ON DELETE CASCADE,
      kind TEXT NOT NULL, raw_value TEXT NOT NULL, normalized_value TEXT NOT NULL, confirmed_at TEXT NOT NULL,
      context TEXT NOT NULL DEFAULT '', reviewer TEXT NOT NULL DEFAULT '', is_new INTEGER NOT NULL DEFAULT 0, UNIQUE(model_id, kind, normalized_value)
    );
    CREATE TABLE IF NOT EXISTS mapping_review_decisions (
      normalized_oe TEXT PRIMARY KEY, raw_oe TEXT NOT NULL, status TEXT NOT NULL,
      model_ids TEXT NOT NULL DEFAULT '[]', note TEXT NOT NULL DEFAULT '', source_file TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS catalog_products (
      normalized_internal TEXT PRIMARY KEY, internal TEXT NOT NULL, common_name TEXT NOT NULL DEFAULT '', vehicle TEXT NOT NULL DEFAULT '',
      oe TEXT NOT NULL DEFAULT '', note TEXT NOT NULL DEFAULT '', image_url TEXT NOT NULL DEFAULT '', image_data TEXT NOT NULL DEFAULT '',
      source_sheet TEXT NOT NULL DEFAULT '', source_row INTEGER NOT NULL DEFAULT 0, source_label TEXT NOT NULL DEFAULT '', updated_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS reviews (
      id INTEGER PRIMARY KEY AUTOINCREMENT, query TEXT NOT NULL, model_id INTEGER, proposed_oe TEXT NOT NULL DEFAULT '[]', proposed_uf TEXT NOT NULL DEFAULT '[]',
      verified_oe TEXT NOT NULL DEFAULT '[]', verified_uf TEXT NOT NULL DEFAULT '[]', platforms TEXT NOT NULL DEFAULT '[]',
      sources TEXT NOT NULL DEFAULT '[]', confidence INTEGER NOT NULL DEFAULT 0, search_provider TEXT NOT NULL DEFAULT '', searched_at TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT '待校对', reviewer TEXT NOT NULL DEFAULT '', reviewed_at TEXT NOT NULL DEFAULT '', note TEXT NOT NULL DEFAULT ''
    );
    CREATE INDEX IF NOT EXISTS idx_reviews_status ON reviews(status);
    CREATE TABLE IF NOT EXISTS versions (
      id INTEGER PRIMARY KEY AUTOINCREMENT, type TEXT NOT NULL, version TEXT NOT NULL, filename TEXT NOT NULL,
      created_at TEXT NOT NULL, operator TEXT NOT NULL, summary TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS quote_jobs (
      id TEXT PRIMARY KEY, filename TEXT NOT NULL, pricing_label TEXT NOT NULL, result_json TEXT NOT NULL, created_at TEXT NOT NULL
    );
  `);
  const ensureColumn = (table, name, definition) => {
    if (!db.prepare(`PRAGMA table_info(${table})`).all().some((column) => column.name === name)) db.exec(`ALTER TABLE ${table} ADD COLUMN ${name} ${definition}`);
  };
  ensureColumn("models", "web_oe_nos", "TEXT NOT NULL DEFAULT '[]'");
  ensureColumn("models", "web_uf_nos", "TEXT NOT NULL DEFAULT '[]'");
  ensureColumn("identifiers", "source_type", "TEXT NOT NULL DEFAULT 'original'");
  ensureColumn("prices", "attribute", "TEXT NOT NULL DEFAULT ''");
  ensureColumn("confirmed_relations", "is_new", "INTEGER NOT NULL DEFAULT 0");
  ensureColumn("reviews", "verified_oe", "TEXT NOT NULL DEFAULT '[]'");
  ensureColumn("reviews", "verified_uf", "TEXT NOT NULL DEFAULT '[]'");
  ensureColumn("reviews", "platforms", "TEXT NOT NULL DEFAULT '[]'");
  migratePricesToAppendOnly();
}

function migratePricesToAppendOnly() {
  const schema = db.prepare("SELECT sql FROM sqlite_master WHERE type='table' AND name='prices'").get()?.sql || "";
  const normalizedSchema = schema.replace(/\s+/g, " ").toUpperCase();
  if (!normalizedSchema.includes("UNIQUE(NORMALIZED_INTERNAL, PACKAGING, TAX, TIER)")) return;
  db.exec(`
    CREATE TABLE prices_append_only (
      id INTEGER PRIMARY KEY AUTOINCREMENT, normalized_internal TEXT NOT NULL, packaging TEXT NOT NULL, tax TEXT NOT NULL, tier TEXT NOT NULL,
      price REAL NOT NULL, attribute TEXT NOT NULL DEFAULT '', source_sheet TEXT NOT NULL, source_column TEXT NOT NULL
    );
    INSERT INTO prices_append_only(id, normalized_internal, packaging, tax, tier, price, attribute, source_sheet, source_column)
      SELECT id, normalized_internal, packaging, tax, tier, price, attribute, source_sheet, source_column FROM prices;
    DROP TABLE prices;
    ALTER TABLE prices_append_only RENAME TO prices;
  `);
}

setupSchema();
const insertIdentifier = db.prepare("INSERT INTO identifiers(model_id, kind, raw_value, normalized_value, source_type) VALUES (?, ?, ?, ?, ?) ON CONFLICT(model_id, kind, normalized_value) DO UPDATE SET raw_value=excluded.raw_value, source_type=excluded.source_type");

function refreshIdentifiers(modelId, model) {
  db.prepare("DELETE FROM identifiers WHERE model_id = ?").run(modelId);
  const webOe = new Set((model.webOeNos || []).map(normalize));
  const webUf = new Set((model.webUfNos || []).map(normalize));
  const groups = [
    ["内部型号", [model.internal, ...model.oldModels], new Set()],
    ["货品编码", model.productCodes, new Set()],
    ["OE号", model.oeNos, webOe],
    ["UF号", model.ufNos, webUf],
  ];
  for (const [kind, values, webValues] of groups) {
    for (const raw of unique(values)) {
      const normalized = normalize(raw);
      if (normalized) insertIdentifier.run(modelId, kind, String(raw), normalized, webValues.has(normalized) ? "web_confirmed" : "original");
    }
  }
}

function upsertModels(models, sourceLabel, sourceType = "original") {
  let added = 0;
  let supplemented = 0;
  let preserved = 0;
  const now = new Date().toISOString();
  db.exec("BEGIN");
  try {
    for (const incoming of models) {
      const key = normalize(incoming.internal);
      const existing = db.prepare("SELECT * FROM models WHERE normalized_internal = ?").get(key);
      if (existing) {
        const merged = {
          internal: existing.internal,
          productCodes: unique([...json(existing.product_codes), ...incoming.productCodes]),
          oldModels: unique([...json(existing.old_models), ...incoming.oldModels]),
          oeNos: unique([...json(existing.oe_nos), ...incoming.oeNos]),
          ufNos: unique([...json(existing.uf_nos), ...incoming.ufNos]),
          webOeNos: json(existing.web_oe_nos),
          webUfNos: json(existing.web_uf_nos),
        };
        const vehicle = mergeTextValues(existing.vehicle, incoming.vehicle);
        const category = mergeTextValues(existing.category, incoming.category);
        const source = mergeTextValues(existing.source_label, sourceLabel);
        const changed = [
          [existing.product_codes, merged.productCodes],
          [existing.old_models, merged.oldModels],
          [existing.oe_nos, merged.oeNos],
          [existing.uf_nos, merged.ufNos],
          [existing.vehicle || "", vehicle],
          [existing.category || "", category],
          [existing.source_label || "", source],
        ].some(([before, after]) => String(before) !== (Array.isArray(after) ? JSON.stringify(after) : String(after)));
        if (changed) {
          db.prepare("UPDATE models SET product_codes=?, old_models=?, oe_nos=?, uf_nos=?, vehicle=?, category=?, source_label=? WHERE id=?").run(
            JSON.stringify(merged.productCodes), JSON.stringify(merged.oldModels), JSON.stringify(merged.oeNos), JSON.stringify(merged.ufNos),
            vehicle, category, source, existing.id,
          );
          refreshIdentifiers(existing.id, { ...merged, vehicle, category });
          supplemented += 1;
        } else {
          preserved += 1;
        }
      } else {
        const result = db.prepare("INSERT INTO models(internal, normalized_internal, product_codes, old_models, oe_nos, uf_nos, vehicle, category, source_type, source_label, captured_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)").run(
          incoming.internal, key, JSON.stringify(incoming.productCodes), JSON.stringify(incoming.oldModels), JSON.stringify(incoming.oeNos), JSON.stringify(incoming.ufNos), incoming.vehicle, incoming.category, sourceType, sourceLabel, now,
        );
        refreshIdentifiers(Number(result.lastInsertRowid), { ...incoming, webOeNos: [], webUfNos: [] });
        added += 1;
      }
    }
    db.exec("COMMIT");
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
  return { added, supplemented, preserved, total: models.length };
}

function upsertPrices(prices) {
  const sameKey = db.prepare("SELECT 1 FROM prices WHERE normalized_internal=? AND packaging=? AND tax=? AND tier=? LIMIT 1");
  const exact = db.prepare("SELECT 1 FROM prices WHERE normalized_internal=? AND packaging=? AND tax=? AND tier=? AND price=? AND attribute=? AND source_sheet=? AND source_column=? LIMIT 1");
  const statement = db.prepare("INSERT INTO prices(normalized_internal, packaging, tax, tier, price, attribute, source_sheet, source_column) VALUES (?, ?, ?, ?, ?, ?, ?, ?)");
  let added = 0;
  let supplemented = 0;
  let preserved = 0;
  db.exec("BEGIN");
  try {
    for (const item of prices) {
      const key = [normalize(item.internal), item.packaging, item.tax, item.tier];
      const values = [...key, item.price, item.attribute || "", item.sourceSheet, item.sourceColumn];
      if (exact.get(...values)) {
        preserved += 1;
        continue;
      }
      const hadKey = Boolean(sameKey.get(...key));
      statement.run(...values);
      if (hadKey) supplemented += 1;
      else added += 1;
    }
    db.exec("COMMIT");
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
  return { added, supplemented, preserved, total: prices.length };
}

function parseCatalog(bufferOrPath) {
  const text = Buffer.isBuffer(bufferOrPath) ? bufferOrPath.toString("utf8") : fs.readFileSync(bufferOrPath, "utf8");
  const payload = JSON.parse(text || "{}");
  return (payload.records || []).map((item) => ({
    internal: String(item.internal || "").trim(),
    normalizedInternal: normalize(item.normalizedInternal || item.internal),
    commonName: String(item.commonName || "").trim(),
    vehicle: String(item.vehicle || "").trim(),
    oe: String(item.oe || "").trim(),
    note: String(item.note || "").trim(),
    imageUrl: String(item.imageUrl || "").trim(),
    imageData: String(item.imageData || "").trim(),
    sourceSheet: String(item.sourceSheet || "").trim(),
    sourceRow: Number(item.sourceRow || 0),
  })).filter((item) => item.normalizedInternal);
}

function upsertCatalog(records, sourceLabel) {
  const statement = db.prepare(`
    INSERT INTO catalog_products(normalized_internal, internal, common_name, vehicle, oe, note, image_url, image_data, source_sheet, source_row, source_label, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(normalized_internal) DO UPDATE SET
      internal=excluded.internal, common_name=excluded.common_name, vehicle=excluded.vehicle, oe=excluded.oe, note=excluded.note,
      image_url=CASE WHEN catalog_products.image_url='' THEN excluded.image_url ELSE catalog_products.image_url END,
      image_data=CASE WHEN catalog_products.image_data='' THEN excluded.image_data ELSE catalog_products.image_data END,
      source_sheet=excluded.source_sheet, source_row=excluded.source_row, source_label=excluded.source_label,
      updated_at=catalog_products.updated_at
  `);
  const existing = new Map(db.prepare("SELECT * FROM catalog_products").all().map((item) => [item.normalized_internal, item]));
  let added = 0;
  let supplemented = 0;
  let preserved = 0;
  const now = new Date().toISOString();
  db.exec("BEGIN");
  try {
    for (const item of records) {
      const previous = existing.get(item.normalizedInternal);
      if (!previous) {
        statement.run(item.normalizedInternal, item.internal, item.commonName, item.vehicle, item.oe, item.note, item.imageUrl, item.imageData, item.sourceSheet, item.sourceRow, sourceLabel, now);
        added += 1;
        continue;
      }
      const merged = {
        internal: mergeTextValues(previous.internal, item.internal),
        commonName: mergeTextValues(previous.common_name, item.commonName),
        vehicle: mergeTextValues(previous.vehicle, item.vehicle),
        oe: mergeTextValues(previous.oe, item.oe),
        note: mergeTextValues(previous.note, item.note),
        imageUrl: previous.image_url || item.imageUrl,
        imageData: previous.image_data || item.imageData,
        sourceSheet: mergeTextValues(previous.source_sheet, item.sourceSheet),
        sourceRow: previous.source_row || item.sourceRow,
        sourceLabel: mergeTextValues(previous.source_label, sourceLabel),
      };
      statement.run(item.normalizedInternal, merged.internal, merged.commonName, merged.vehicle, merged.oe, merged.note, merged.imageUrl, merged.imageData, merged.sourceSheet, merged.sourceRow, merged.sourceLabel, now);
      const changed = [
        [previous.internal, merged.internal], [previous.common_name, merged.commonName], [previous.vehicle, merged.vehicle], [previous.oe, merged.oe],
        [previous.note, merged.note], [previous.source_sheet, merged.sourceSheet], [previous.source_row, merged.sourceRow], [previous.source_label, merged.sourceLabel],
      ].some(([before, after]) => String(before || "") !== String(after || ""));
      if (changed) supplemented += 1;
      else preserved += 1;
      existing.set(item.normalizedInternal, {
        ...previous,
        internal: merged.internal, common_name: merged.commonName, vehicle: merged.vehicle, oe: merged.oe, note: merged.note,
        image_url: merged.imageUrl, image_data: merged.imageData, source_sheet: merged.sourceSheet, source_row: merged.sourceRow, source_label: merged.sourceLabel,
      });
    }
    db.exec("COMMIT");
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
  return { added, supplemented, preserved, total: records.length, images: records.filter((item) => item.imageUrl || item.imageData).length };
}

function catalogForInternal(internal) {
  const row = db.prepare("SELECT * FROM catalog_products WHERE normalized_internal=?").get(normalize(internal));
  if (!row) return null;
  return {
    internal: row.internal,
    commonName: row.common_name,
    vehicle: row.vehicle,
    image: row.image_data || row.image_url,
    sourceSheet: row.source_sheet,
  };
}

function insertVersion(type, filename, operator, summary) {
  const count = db.prepare("SELECT COUNT(*) AS count FROM versions WHERE type=?").get(type).count + 1;
  const prefix = type === "型号库" ? "M" : type === "价格库" ? "P" : "C";
  const version = `${prefix}${String(count).padStart(3, "0")}`;
  db.prepare("INSERT INTO versions(type, version, filename, created_at, operator, summary) VALUES (?, ?, ?, ?, ?, ?)").run(type, version, filename, new Date().toISOString(), operator, JSON.stringify(summary));
  return version;
}

function seed() {
  if (db.prepare("SELECT value FROM meta WHERE key='seeded'").get()) return;
  // 线上多人版会先从持久数据库恢复；部署包没有初始化源表时不应在恢复前报错。
  if (!fs.existsSync(sourceModelPath) || !fs.existsSync(sourcePricePath)) return;
  const modelSummary = upsertModels(parseModels(sourceModelPath), "大厂型号匹配260809.xlsx");
  const priceSummary = upsertPrices(parsePrices(sourcePricePath));
  insertVersion("型号库", "大厂型号匹配260809.xlsx", "系统初始化", modelSummary);
  insertVersion("价格库", "外贸价格体系60319.xlsx", "系统初始化", priceSummary);
  db.prepare("INSERT INTO meta(key,value) VALUES('seeded',?)").run(new Date().toISOString());
}

function seedCatalog() {
  if (!fs.existsSync(sourceCatalogPath) || db.prepare("SELECT value FROM meta WHERE key='catalog_seeded_v1'").get()) return;
  const summary = upsertCatalog(parseCatalog(sourceCatalogPath), "海福来点火线圈目录（带通用名）60625.xlsx");
  insertVersion("产品目录", "海福来点火线圈目录（带通用名）60625.xlsx", "系统初始化", summary);
  db.prepare("INSERT INTO meta(key,value) VALUES('catalog_seeded_v1',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value").run(new Date().toISOString());
}

function migrateIdentifierParser() {
  if (db.prepare("SELECT value FROM meta WHERE key='mapping_parser_version'").get()?.value === "3") return;
  const rows = db.prepare("SELECT * FROM models").all();
  db.exec("BEGIN");
  try {
    for (const row of rows) {
      const model = {
        internal: row.internal,
        productCodes: json(row.product_codes),
        oldModels: json(row.old_models),
        oeNos: parseIdentifierList(json(row.oe_nos).join("\n")).filter((item) => !/^UF/i.test(item)),
        ufNos: unique([
          ...parseIdentifierList(json(row.uf_nos).join("\n")),
          ...parseIdentifierList(json(row.oe_nos).join("\n")).filter((item) => /^UF/i.test(item)),
        ]),
        webOeNos: json(row.web_oe_nos),
        webUfNos: json(row.web_uf_nos),
      };
      db.prepare("UPDATE models SET oe_nos=?, uf_nos=? WHERE id=?").run(JSON.stringify(model.oeNos), JSON.stringify(model.ufNos), row.id);
      refreshIdentifiers(row.id, model);
    }
      db.prepare("INSERT INTO meta(key,value) VALUES('mapping_parser_version','3') ON CONFLICT(key) DO UPDATE SET value='3'").run();
    db.exec("COMMIT");
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
}

function hydratePriceAttributes() {
  if (!fs.existsSync(sourcePricePath)) return;
  const attributes = new Map();
  for (const item of parsePrices(sourcePricePath)) {
    if (item.attribute) attributes.set(normalize(item.internal), item.attribute);
  }
  const update = db.prepare("UPDATE prices SET attribute=? WHERE normalized_internal=? AND (attribute='' OR attribute IS NULL)");
  for (const [internal, attribute] of attributes) update.run(attribute, internal);
}

function migrateConfirmedReviewMappings() {
  if (db.prepare("SELECT value FROM meta WHERE key='confirmed_review_sync_version'").get()?.value === "3") return { changed: 0 };
  const reviews = db.prepare("SELECT * FROM reviews WHERE status='已确认' AND model_id IS NOT NULL ORDER BY reviewed_at, id").all();
  if (!reviews.length) return { changed: 0 };
  let changed = 0;
  for (const review of reviews) {
    const row = db.prepare("SELECT * FROM models WHERE id=?").get(review.model_id);
    if (!row) continue;
    const inputs = unique([review.query, ...json(review.proposed_oe), ...json(review.proposed_uf)]).join(" / ");
    const relation = addConfirmedInput({ ...row, __created: false }, inputs, "历史确认同步", review.reviewer || "当前用户");
    changed += relation.addedOldModels.length + relation.addedOeNos.length + relation.addedUfNos.length + relation.normalizedLegacy;
  }
  if (changed) insertVersion("型号库", "大厂型号匹配（历史确认同步）.xlsx", "系统迁移", {
    added: 0, updated: reviews.length, relationsAdded: changed,
    total: db.prepare("SELECT COUNT(*) AS count FROM models").get().count,
    context: "历史确认同步",
  });
  db.prepare("INSERT INTO meta(key,value) VALUES('confirmed_review_sync_version','3') ON CONFLICT(key) DO UPDATE SET value='3'").run();
  return { changed };
}

export function initKnowledgeBase() {
  setupSchema();
  seed();
  seedCatalog();
  migrateIdentifierParser();
  hydratePriceAttributes();
  migrateConfirmedReviewMappings();
}

const persistedTables = ["meta", "models", "mapping_review_decisions", "identifiers", "confirmed_relations", "prices", "catalog_products", "reviews", "versions", "quote_jobs"];
const deleteOrder = ["mapping_review_decisions", "identifiers", "confirmed_relations", "reviews", "quote_jobs", "prices", "catalog_products", "versions", "models", "meta"];
const insertOrder = ["meta", "models", "mapping_review_decisions", "identifiers", "confirmed_relations", "prices", "catalog_products", "reviews", "versions", "quote_jobs"];
let persistentSql;
let persistentReady;
let localRevision = 0;
let localStateQueue = Promise.resolve();

function getPersistentSql() {
  if (!process.env.DATABASE_URL) return null;
  persistentSql ||= neon(process.env.DATABASE_URL);
  return persistentSql;
}

function encodeSnapshot() {
  const snapshot = {};
  for (const table of persistedTables) {
    snapshot[table] = table === "quote_jobs"
      ? db.prepare("SELECT * FROM quote_jobs ORDER BY created_at DESC LIMIT 20").all()
      : db.prepare(`SELECT * FROM ${table}`).all();
  }
  return gzipSync(Buffer.from(JSON.stringify(snapshot))).toString("base64");
}

function restoreSnapshot(payload) {
  const snapshot = JSON.parse(gunzipSync(Buffer.from(payload, "base64")).toString("utf8"));
  db.exec("BEGIN");
  try {
    for (const table of deleteOrder) db.prepare(`DELETE FROM ${table}`).run();
    for (const table of insertOrder) {
      for (const row of snapshot[table] || []) {
        const columns = Object.keys(row);
        if (!columns.length) continue;
        const statement = db.prepare(`INSERT INTO ${table} (${columns.map((column) => `"${column}"`).join(",")}) VALUES (${columns.map(() => "?").join(",")})`);
        statement.run(...columns.map((column) => row[column]));
      }
    }
    db.exec("COMMIT");
    hydratePriceAttributes();
    migrateConfirmedReviewMappings();
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
}

const delay = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds));

function queueLocalState(task) {
  const result = localStateQueue.then(task, task);
  localStateQueue = result.catch(() => {});
  return result;
}

async function ensurePersistentState(sql) {
  persistentReady ||= (async () => {
    await sql`CREATE TABLE IF NOT EXISTS hifly_app_state (
      id INTEGER PRIMARY KEY,
      payload TEXT NOT NULL,
      revision BIGINT NOT NULL DEFAULT 1,
      lock_token TEXT,
      lock_expires_at TIMESTAMPTZ,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )`;
    const existing = await sql`SELECT id FROM hifly_app_state WHERE id=1`;
    if (!existing.length) await sql`INSERT INTO hifly_app_state (id, payload) VALUES (1, ${encodeSnapshot()}) ON CONFLICT (id) DO NOTHING`;
  })();
  return persistentReady;
}

export async function withPersistentState(callback, { write = false } = {}) {
  const sql = getPersistentSql();
  if (!sql) return callback();
  await ensurePersistentState(sql);
  return queueLocalState(async () => {
    if (!write) {
      const [remote] = await sql`SELECT revision FROM hifly_app_state WHERE id=1`;
      if (Number(remote.revision) !== localRevision) {
        const [state] = await sql`SELECT payload, revision FROM hifly_app_state WHERE id=1`;
        restoreSnapshot(state.payload);
        localRevision = Number(state.revision);
      }
      return callback();
    }

    const token = randomUUID();
    let locked;
    for (let attempt = 0; attempt < 40; attempt += 1) {
      const rows = await sql`UPDATE hifly_app_state
        SET lock_token=${token}, lock_expires_at=NOW() + INTERVAL '75 seconds'
        WHERE id=1 AND (lock_token IS NULL OR lock_expires_at < NOW())
        RETURNING revision`;
      if (rows.length) { locked = rows[0]; break; }
      await delay(250);
    }
    if (!locked) throw new Error("共享数据正在被其他用户更新，请稍后重试");

    try {
      if (Number(locked.revision) !== localRevision) {
        const [state] = await sql`SELECT payload, revision FROM hifly_app_state WHERE id=1`;
        restoreSnapshot(state.payload);
        localRevision = Number(state.revision);
      }
      const result = await callback();
      const saved = await sql`UPDATE hifly_app_state
        SET payload=${encodeSnapshot()}, revision=revision+1, updated_at=NOW(), lock_token=NULL, lock_expires_at=NULL
        WHERE id=1 AND lock_token=${token} RETURNING revision`;
      if (!saved.length) throw new Error("共享数据保存冲突，请重新操作");
      localRevision = Number(saved[0].revision);
      return result;
    } catch (error) {
      await sql`UPDATE hifly_app_state SET lock_token=NULL, lock_expires_at=NULL WHERE id=1 AND lock_token=${token}`.catch(() => {});
      throw error;
    }
  });
}

const VEHICLE_TRANSLATIONS = [
  ["梅赛德斯奔驰", "Mercedes-Benz"], ["奔驰", "Mercedes-Benz"], ["克莱斯勒", "Chrysler"], ["雪佛兰", "Chevrolet"], ["凯迪拉克", "Cadillac"],
  ["丰田大发", "Toyota Daihatsu"], ["丰田", "Toyota"], ["大发", "Daihatsu"], ["本田", "Honda"], ["讴歌", "Acura"], ["日产", "Nissan"], ["尼桑", "Nissan"],
  ["马自达", "Mazda"], ["雷克萨斯", "Lexus"], ["凌志", "Lexus"], ["三菱", "Mitsubishi"], ["斯巴鲁", "Subaru"], ["铃木", "Suzuki"],
  ["现代", "Hyundai"], ["起亚", "Kia"], ["福特", "Ford"], ["雪铁龙", "Citroen"], ["标致", "Peugeot"], ["雷诺", "Renault"], ["欧宝", "Opel"],
  ["宝马", "BMW"], ["保时捷", "Porsche"], ["沃尔沃", "Volvo"], ["大众", "Volkswagen"], ["捷豹", "Jaguar"], ["路虎", "Land Rover"],
  ["道奇", "Dodge"], ["吉普", "Jeep"], ["Jeep", "Jeep"], ["五十铃", "Isuzu"], ["五菱", "Wuling"], ["比亚迪", "BYD"], ["吉利", "Geely"],
  ["奇瑞", "Chery"], ["哈弗", "Haval"], ["长城", "Great Wall"], ["众泰", "Zotye"], ["传祺", "Trumpchi"], ["名爵", "MG"], ["荣威", "Roewe"],
  ["一汽", "FAW"], ["上汽", "SAIC"], ["广汽", "GAC"], ["东风", "Dongfeng"], ["东南", "Soueast"], ["北汽", "BAIC"], ["江淮", "JAC"],
  ["华晨", "Brilliance"], ["中华", "Brilliance"], ["猎豹", "Cheetah"], ["力帆", "Lifan"], ["海马", "Haima"], ["纳智捷", "Luxgen"], ["陆风", "Landwind"], ["长安", "Changan"], ["奔腾", "Bestune"], ["青年莲花", "Young Lotus"], ["三星", "Samsung"], ["高尔", "Golf"],
  ["雅阁", "Accord"], ["思域", "Civic"], ["锋范", "City"], ["CRV", "CR-V"], ["凌派", "Crider"], ["智跑", "Sportage"], ["思铂睿", "Spirior"], ["飞度", "Fit"],
  ["雅力士", "Yaris"], ["威驰", "Vios"], ["卡罗拉", "Corolla"], ["花冠", "Corolla"], ["凯美瑞", "Camry"], ["佳美", "Camry"], ["汉兰达", "Highlander"], ["皇冠", "Crown"], ["锐志", "Mark X"], ["塞纳", "Sienna"], ["普锐斯", "Prius"], ["陆地巡洋舰", "Land Cruiser"], ["霸道", "Prado"], ["大发特锐", "Daihatsu Terios"], ["特锐", "Terios"], ["雷凌", "Levin"],
  ["欧蓝德", "Outlander"], ["帕杰罗", "Pajero"], ["蒙特罗", "Montero"], ["蓝瑟", "Lancer"], ["劲炫", "ASX"], ["伊柯丽斯", "Eclipse"], ["速跑", "Sport"], ["雨燕", "Swift"], ["天语", "SX4"], ["马六", "Mazda6"], ["昂克赛拉", "Mazda3 Axela"], ["玛驰", "March"], ["轩逸", "Sylphy"], ["骐达", "Tiida"], ["天籁", "Teana"], ["逍客", "Qashqai"], ["奇骏", "X-Trail"], ["风度", "Cefiro"], ["阳光", "Sunny"],
  ["嘉年华", "Fiesta"], ["福克斯", "Focus"], ["蒙迪欧", "Mondeo"], ["翼虎", "Kuga"], ["锐界", "Edge"], ["探险者", "Explorer"], ["神探", "Explorer"], ["猎鹰", "Falcon"], ["领土", "Territory"], ["科鲁兹", "Cruze"], ["赛欧", "Sail"], ["迈锐宝", "Malibu"], ["科帕奇", "Captiva"], ["索罗德", "Silverado"], ["英帕拉", "Impala"], ["索尼克", "Sonic"], ["科迈罗", "Camaro"], ["拉赛蒂", "Lacetti"], ["英朗", "Excelle"], ["君威", "Regal"], ["君越", "LaCrosse"], ["昂科拉", "Encore"], ["伊兰特", "Elantra"], ["名图", "Mistra"],
  ["自由光", "Cherokee"], ["指南者", "Compass"], ["自由客", "Patriot"], ["酷威", "Journey"], ["锋哲", "Avenger"], ["铂锐", "Sebring"], ["森林人", "Forester"], ["力狮", "Legacy"], ["翼豹", "Impreza"], ["帝豪", "Emgrand"], ["博越", "Boyue"], ["缤瑞", "Binrui"], ["缤越", "Binyue"], ["远景", "Vision"], ["瑞虎", "Tiggo"], ["风云", "Fengyun"], ["影豹", "Empow"], ["大迈", "Damai"], ["风骏", "Wingle"], ["大狗", "Big Dog"], ["宏光", "Hongguang"], ["荣光", "Rongguang"], ["凯捷", "Kaijie"], ["征程", "Journey"], ["杨光", "Sunshine"], ["森雅", "Xenia"], ["景逸", "Joyear"], ["菱智", "Lingzhi"], ["风光", "Fengon"], ["风神", "Aeolus"], ["骏捷", "Junjie"], ["尊驰", "Zunchi"],
  ["发动机", "engine"], ["天然气", "CNG"], ["混合动力", "Hybrid"], ["手动", "Manual"], ["两厢", "Hatchback"], ["三厢", "Sedan"], ["进口", "Imported"], ["国内", "Domestic"], ["外贸", "Export"], ["日式", "Japanese style"], ["法式", "French style"], ["专用", "Dedicated"], ["红色", "Red"], ["黑色", "Black"], ["白色", "White"], ["蓝色", "Blue"], ["红插", "Red connector"], ["白插", "White connector"], ["黑插", "Black connector"], ["黑胶面", "Black rubber surface"], ["白胶面", "White rubber surface"], ["两插", "2-pin"], ["三插", "3-pin"], ["四孔", "4-hole"], ["三孔", "3-hole"], ["二针", "2-pin"], ["三针", "3-pin"], ["长杆", "Long stem"], ["短杆", "Short stem"], ["短", "Short"], ["前排", "Front row"], ["后排", "Rear row"], ["宽卡", "Wide clip"], ["系统", "system"], ["总成", "assembly"], ["新参数", "New spec"],
];

function vehicleToEnglish(value = "") {
  let result = String(value || "").trim();
  if (!result) return "";
  for (const [source, target] of [...VEHICLE_TRANSLATIONS].sort((a, b) => b[0].length - a[0].length)) result = result.split(source).join(target);
  result = result.replace(/(\d{4})年后/g, "after $1").replace(/(\d{2,4})款/g, "$1 model");
  result = result.replace(/[（]/g, " (").replace(/[）]/g, ")").replace(/[、，]/g, ", ").replace(/[；;]/g, "; ").replace(/[：:]/g, ": ");
  return result.replace(/\s+/g, " ").replace(/\(\s+/g, "(").replace(/\s+\)/g, ")").trim();
}

function serializeModel(row, matchedKind = "") {
  const rawOeNos = json(row.oe_nos);
  const rawUfNos = json(row.uf_nos);
  const rawWebOeNos = json(row.web_oe_nos);
  const rawWebUfNos = json(row.web_uf_nos);
  const displayInternal = normalizeRedShellInternal(row.internal, catalogForInternal(row.internal));
  return {
    id: row.id, internal: row.internal, displayInternal, redShellAdjusted: displayInternal !== row.internal, productCodes: json(row.product_codes), oldModels: json(row.old_models),
    oeNos: rawOeNos.filter((item) => !/^UF/i.test(String(item).trim())),
    ufNos: unique([...rawUfNos, ...rawOeNos.filter((item) => /^UF/i.test(String(item).trim()))]), vehicle: row.vehicle, category: row.category,
    vehicleEnglish: vehicleToEnglish(row.vehicle),
    webOeNos: rawWebOeNos.filter((item) => !/^UF/i.test(String(item).trim())),
    webUfNos: unique([...rawWebUfNos, ...rawWebOeNos.filter((item) => /^UF/i.test(String(item).trim()))]),
    sourceType: row.source_type, sourceLabel: row.source_label, sourceUrl: row.source_url, capturedAt: row.captured_at,
    confidence: row.confidence, reviewStatus: row.review_status, matchedKind,
  };
}

function findLocalExact(query) {
  const key = normalize(query);
  const exact = db.prepare("SELECT m.*, i.kind AS matched_kind FROM identifiers i JOIN models m ON m.id=i.model_id WHERE i.normalized_value=? ORDER BY CASE m.source_type WHEN 'original' THEN 0 ELSE 1 END, m.id LIMIT 12").all(key);
  const serialized = [...new Map(exact.map((row) => [row.id, serializeModel(row, row.matched_kind)])).values()];
  if (serialized.some((item) => item.matchedKind === "内部型号" && normalize(item.internal) === key)) return serialized.filter((item) => item.matchedKind === "内部型号" && normalize(item.internal) === key);
  const latestChoice = db.prepare("SELECT model_id FROM confirmed_relations WHERE normalized_value=? ORDER BY confirmed_at DESC, id DESC LIMIT 1").get(key);
  return latestChoice && serialized.some((item) => item.id === latestChoice.model_id) ? serialized.filter((item) => item.id === latestChoice.model_id) : serialized;
}

function findLocal(query) {
  const key = normalize(query);
  const exact = findLocalExact(query);
  if (exact.length) return exact;
  if (key.length < 4) return [];
  const rows = db.prepare("SELECT m.*, i.kind AS matched_kind FROM identifiers i JOIN models m ON m.id=i.model_id WHERE i.normalized_value LIKE ? ORDER BY m.id LIMIT 12").all(`${key}%`);
  return [...new Map(rows.map((row) => [row.id, serializeModel(row, row.matched_kind)])).values()];
}

function findUniversalLocal(query, { exact = false } = {}) {
  const aliases = splitUniversalInputs(query);
  const universal = aliases.length > 1;
  if (!universal) return { aliases: aliases.length ? aliases : [String(query || "").trim()], universal, models: exact ? findLocalExact(query) : findLocal(query) };
  const collected = new Map();
  for (const alias of aliases) {
    const matches = exact ? findLocalExact(alias) : findLocal(alias);
    for (const model of matches) {
      const existing = collected.get(model.id);
      if (existing) existing.matchedAliases = unique([...(existing.matchedAliases || []), alias]);
      else collected.set(model.id, { ...model, matchedAliases: [alias] });
    }
  }
  return { aliases, universal, models: [...collected.values()] };
}

function ensureUserModel(targetInternal) {
  const value = String(targetInternal || "").trim().toUpperCase();
  const key = normalize(value);
  if (!value || key.length < 4 || !/\d/.test(key)) throw new Error("请填写有效的内部型号");
  const existing = db.prepare("SELECT * FROM models WHERE normalized_internal=?").get(key);
  if (existing) return { ...existing, __created: false };
  const now = new Date().toISOString();
  const result = db.prepare("INSERT INTO models(internal, normalized_internal, product_codes, old_models, oe_nos, uf_nos, vehicle, category, source_type, source_label, captured_at, confidence, review_status) VALUES (?, ?, '[]', '[]', '[]', '[]', '', '', 'user_confirmed', '用户确认新增', ?, 100, '已确认')").run(value, key, now);
  const row = db.prepare("SELECT * FROM models WHERE id=?").get(Number(result.lastInsertRowid));
  refreshIdentifiers(row.id, { internal: row.internal, productCodes: [], oldModels: [], oeNos: [], ufNos: [], webOeNos: [], webUfNos: [] });
  return { ...row, __created: true };
}

function classifyConfirmedInput(input, internal) {
  const relation = { oeNos: [], ufNos: [], oldModels: [] };
  for (const item of splitUniversalInputs(input)) {
    const value = String(item || "").trim().toUpperCase();
    if (!value || normalize(value) === normalize(internal)) continue;
    if (/^DQ/i.test(value)) relation.oldModels.push(value);
    else if (/^UF/i.test(value)) relation.ufNos.push(value);
    else relation.oeNos.push(value);
  }
  relation.oeNos = unique(relation.oeNos);
  relation.ufNos = unique(relation.ufNos);
  relation.oldModels = unique(relation.oldModels);
  return relation;
}

function addConfirmedInput(row, input, context = "型号查询", reviewer = "当前用户", extras = {}) {
  const relation = classifyConfirmedInput(input, row.internal);
  relation.oeNos = unique([...relation.oeNos, ...(extras.oeNos || [])]);
  relation.ufNos = unique([...relation.ufNos, ...(extras.ufNos || [])]);
  const beforeOld = new Set(json(row.old_models).map(normalize));
  const beforeOe = new Set(json(row.oe_nos).map(normalize));
  const beforeUf = new Set(json(row.uf_nos).map(normalize));
  const expandedOldModels = expandUniversalValues(json(row.old_models));
  const expandedOeNos = expandUniversalValues(json(row.oe_nos)).filter((value) => !/^UF/i.test(value));
  const expandedUfNos = unique([...expandUniversalValues(json(row.uf_nos)), ...expandUniversalValues(json(row.oe_nos)).filter((value) => /^UF/i.test(value))]);
  const normalizedLegacy = [...json(row.old_models), ...json(row.oe_nos), ...json(row.uf_nos)].filter((value) => splitUniversalInputs(value).length > 1).length;
  const addedOldModels = relation.oldModels.filter((value) => !beforeOld.has(normalize(value)));
  const addedOeNos = relation.oeNos.filter((value) => !beforeOe.has(normalize(value)));
  const addedUfNos = relation.ufNos.filter((value) => !beforeUf.has(normalize(value)));
  const model = {
    internal: row.internal,
    productCodes: json(row.product_codes),
    oldModels: unique([...expandedOldModels, ...relation.oldModels]),
    oeNos: unique([...expandedOeNos, ...relation.oeNos]),
    ufNos: unique([...expandedUfNos, ...relation.ufNos]),
    webOeNos: expandUniversalValues(json(row.web_oe_nos)).filter((value) => !/^UF/i.test(value)),
    webUfNos: unique([...expandUniversalValues(json(row.web_uf_nos)), ...expandUniversalValues(json(row.web_oe_nos)).filter((value) => /^UF/i.test(value))]),
  };
  model.webOeNos = unique([...model.webOeNos, ...(extras.webOeNos || []).filter((value) => model.oeNos.some((item) => normalize(item) === normalize(value)))]);
  model.webUfNos = unique([...model.webUfNos, ...(extras.webUfNos || []).filter((value) => model.ufNos.some((item) => normalize(item) === normalize(value)))]);
  const now = new Date().toISOString();
  db.prepare("UPDATE models SET old_models=?, oe_nos=?, uf_nos=?, web_oe_nos=?, web_uf_nos=?, captured_at=? WHERE id=?").run(JSON.stringify(model.oldModels), JSON.stringify(model.oeNos), JSON.stringify(model.ufNos), JSON.stringify(model.webOeNos), JSON.stringify(model.webUfNos), now, row.id);
  refreshIdentifiers(row.id, model);
  const insertConfirmed = db.prepare("INSERT INTO confirmed_relations(model_id, kind, raw_value, normalized_value, confirmed_at, context, reviewer, is_new) VALUES (?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT(model_id, kind, normalized_value) DO UPDATE SET raw_value=excluded.raw_value, confirmed_at=excluded.confirmed_at, context=excluded.context, reviewer=excluded.reviewer, is_new=MAX(confirmed_relations.is_new, excluded.is_new)");
  if (row.__created) insertConfirmed.run(row.id, "内部型号", row.internal, normalize(row.internal), now, context, reviewer, 1);
  for (const value of relation.oldModels) insertConfirmed.run(row.id, "老型号", value, normalize(value), now, context, reviewer, addedOldModels.some((item) => normalize(item) === normalize(value)) ? 1 : 0);
  for (const value of relation.oeNos) insertConfirmed.run(row.id, "OE号", value, normalize(value), now, context, reviewer, addedOeNos.some((item) => normalize(item) === normalize(value)) ? 1 : 0);
  for (const value of relation.ufNos) insertConfirmed.run(row.id, "UF号", value, normalize(value), now, context, reviewer, addedUfNos.some((item) => normalize(item) === normalize(value)) ? 1 : 0);
  return { ...relation, addedOldModels, addedOeNos, addedUfNos, normalizedLegacy, createdModel: Boolean(row.__created), changed: Boolean(row.__created || addedOldModels.length || addedOeNos.length || addedUfNos.length || normalizedLegacy) };
}

function resolveReviewModel(code) {
  const key = normalize(code);
  if (!key) return null;
  const exact = db.prepare("SELECT * FROM models WHERE normalized_internal=?").get(key);
  if (exact) return exact;
  const aliases = db.prepare("SELECT DISTINCT m.* FROM identifiers i JOIN models m ON m.id=i.model_id WHERE i.normalized_value=? ORDER BY m.id").all(key);
  if (aliases.length === 1) return aliases[0];
  const prefixes = db.prepare("SELECT * FROM models WHERE normalized_internal LIKE ? ORDER BY LENGTH(normalized_internal), id").all(`${key}%`);
  return prefixes.length === 1 ? prefixes[0] : null;
}

function savedMappingDecision(normalizedOe) {
  const row = db.prepare("SELECT * FROM mapping_review_decisions WHERE normalized_oe=?").get(normalizedOe);
  if (!row) return null;
  return { ...row, modelIds: json(row.model_ids).map(Number).filter(Number.isFinite) };
}

function prepareMappingReview(parsed, filename) {
  const decisionMap = new Map();
  for (const item of parsed.decisions || []) {
    for (const oe of item.oeNos || []) {
      const normalizedOe = normalize(oe);
      if (!normalizedOe) continue;
      const note = String(item.note || "").trim();
      const blockedByNote = /无法确认|无法查询/.test(note);
      const directRows = unique(item.confirmModels || []).map(resolveReviewModel).filter(Boolean);
      const fallbackRows = directRows.length ? [] : reviewModelCodes(note).map(resolveReviewModel).filter(Boolean);
      const rows = [...new Map([...directRows, ...fallbackRows].map((row) => [row.id, row])).values()];
      const unresolvedCodes = unique((directRows.length ? item.confirmModels : reviewModelCodes(note)).filter((code) => !resolveReviewModel(code)));
      const status = blockedByNote ? "blocked" : rows.length ? "confirmed" : "unresolved";
      decisionMap.set(normalizedOe, {
        normalizedOe,
        rawOe: oe,
        status,
        modelIds: rows.map((row) => row.id),
        modelCodes: rows.map((row) => row.internal),
        unresolvedCodes,
        note,
        sourceFile: filename,
        authoritative: true,
      });
    }
  }

  const relationPairs = new Map();
  const unresolvedCodes = new Set();
  const addPair = (oe, modelId) => {
    const key = `${modelId}|${normalize(oe)}`;
    if (modelId && normalize(oe)) relationPairs.set(key, { oe, modelId });
  };
  for (const decision of decisionMap.values()) {
    if (decision.status === "confirmed") decision.modelIds.forEach((modelId) => addPair(decision.rawOe, modelId));
    decision.unresolvedCodes.forEach((code) => unresolvedCodes.add(code));
  }
  for (const item of parsed.mergedRelations || []) {
    for (const oe of item.oeNos || []) {
      const normalizedOe = normalize(oe);
      const override = decisionMap.get(normalizedOe) || savedMappingDecision(normalizedOe);
      if (override) {
        if (override.status === "confirmed") override.modelIds.forEach((modelId) => addPair(oe, modelId));
        continue;
      }
      for (const code of item.modelCodes || []) {
        const row = resolveReviewModel(code);
        if (row) addPair(oe, row.id);
        else unresolvedCodes.add(code);
      }
    }
  }

  const currentPairs = new Set(db.prepare("SELECT model_id, normalized_value FROM identifiers WHERE kind='OE号'").all().map((row) => `${row.model_id}|${row.normalized_value}`));
  const added = [...relationPairs.keys()].filter((key) => !currentPairs.has(key)).length;
  let removed = 0;
  for (const decision of decisionMap.values()) {
    if (decision.status !== "confirmed") continue;
    const selected = new Set(decision.modelIds);
    removed += db.prepare("SELECT DISTINCT model_id FROM identifiers WHERE kind='OE号' AND normalized_value=?").all(decision.normalizedOe).filter((row) => !selected.has(row.model_id)).length;
  }
  const incomingOes = new Set([...(parsed.decisions || []).flatMap((item) => item.oeNos || []), ...(parsed.mergedRelations || []).flatMap((item) => item.oeNos || []).map((value) => String(value))].map(normalize).filter(Boolean));
  const confirmed = [...decisionMap.values()].filter((item) => item.status === "confirmed").length;
  const blocked = [...decisionMap.values()].filter((item) => item.status === "blocked").length;
  const unresolved = [...decisionMap.values()].filter((item) => item.status === "unresolved").length;
  const preserved = relationPairs.size - added;
  return {
    decisions: [...decisionMap.values()],
    relationItems: [...relationPairs.values()],
    summary: {
      incoming: incomingOes.size,
      existing: db.prepare("SELECT COUNT(*) AS count FROM models").get().count,
      added,
      supplemented: confirmed,
      preserved: Math.max(0, preserved) + blocked,
      conflicts: blocked + unresolved + unresolvedCodes.size,
      confirmed,
      blocked,
      unresolved: unresolved + unresolvedCodes.size,
      relations: relationPairs.size,
      removed,
      unresolvedModels: [...unresolvedCodes],
      filename,
    },
  };
}

function removeOeRelation(modelId, normalizedOe) {
  const row = db.prepare("SELECT * FROM models WHERE id=?").get(modelId);
  if (!row) return false;
  const key = normalize(normalizedOe);
  const filter = (values) => json(values).filter((value) => normalize(value) !== key);
  const nextOe = filter(row.oe_nos);
  const nextUf = filter(row.uf_nos);
  const nextWebOe = filter(row.web_oe_nos);
  const nextWebUf = filter(row.web_uf_nos);
  const changed = [nextOe.length !== json(row.oe_nos).length, nextUf.length !== json(row.uf_nos).length, nextWebOe.length !== json(row.web_oe_nos).length, nextWebUf.length !== json(row.web_uf_nos).length].some(Boolean);
  if (!changed) return false;
  db.prepare("UPDATE models SET oe_nos=?, uf_nos=?, web_oe_nos=?, web_uf_nos=?, captured_at=? WHERE id=?").run(JSON.stringify(nextOe), JSON.stringify(nextUf), JSON.stringify(nextWebOe), JSON.stringify(nextWebUf), new Date().toISOString(), modelId);
  refreshIdentifiers(modelId, { internal: row.internal, productCodes: json(row.product_codes), oldModels: json(row.old_models), oeNos: nextOe, ufNos: nextUf, webOeNos: nextWebOe, webUfNos: nextWebUf });
  db.prepare("DELETE FROM confirmed_relations WHERE model_id=? AND kind='OE号' AND normalized_value=?").run(modelId, key);
  return true;
}

function applyMappingReviewParsed(parsed, filename, operator) {
  const prepared = prepareMappingReview(parsed, filename);
  let added = 0;
  let removed = 0;
  let confirmed = 0;
  let blocked = 0;
  let unresolved = 0;
  let relationsAdded = 0;
  db.exec("BEGIN");
  try {
    const saveDecision = db.prepare("INSERT INTO mapping_review_decisions(normalized_oe, raw_oe, status, model_ids, note, source_file, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?) ON CONFLICT(normalized_oe) DO UPDATE SET raw_oe=excluded.raw_oe, status=excluded.status, model_ids=excluded.model_ids, note=excluded.note, source_file=excluded.source_file, updated_at=excluded.updated_at");
    for (const decision of prepared.decisions) {
      saveDecision.run(decision.normalizedOe, decision.rawOe, decision.status, JSON.stringify(decision.modelIds), decision.note, filename, new Date().toISOString());
      if (decision.status === "confirmed") {
        confirmed += 1;
        const selected = new Set(decision.modelIds);
        const current = db.prepare("SELECT DISTINCT model_id FROM identifiers WHERE kind='OE号' AND normalized_value=?").all(decision.normalizedOe);
        for (const row of current) if (!selected.has(row.model_id) && removeOeRelation(row.model_id, decision.normalizedOe)) removed += 1;
      } else if (decision.status === "blocked") blocked += 1;
      else unresolved += 1;
    }
    for (const item of prepared.relationItems) {
      const row = db.prepare("SELECT * FROM models WHERE id=?").get(item.modelId);
      if (!row) continue;
      const relation = addConfirmedInput({ ...row, __created: false }, item.oe, "表格关系核验", operator, { oeNos: [item.oe] });
      added += relation.addedOeNos.length;
      relationsAdded += relation.addedOldModels.length + relation.addedOeNos.length + relation.addedUfNos.length;
    }
    const summary = { ...prepared.summary, added, removed, confirmed, blocked, unresolved, relationsAdded, total: db.prepare("SELECT COUNT(*) AS count FROM models").get().count };
    const version = insertVersion("型号库", filename, operator, summary);
    db.exec("COMMIT");
    return { summary, version };
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
}

function recordUserConfirmation(query, modelId, relation, context, reviewer = "当前用户") {
  const now = new Date().toISOString();
  db.prepare("INSERT INTO reviews(query, model_id, proposed_oe, proposed_uf, verified_oe, verified_uf, platforms, sources, confidence, search_provider, searched_at, status, reviewer, reviewed_at, note) VALUES (?, ?, ?, ?, ?, ?, '[]', '[]', 100, ?, ?, '已确认', ?, ?, ?)").run(
    String(query || "").trim(), modelId, JSON.stringify(relation.oeNos), JSON.stringify(relation.ufNos), JSON.stringify(relation.oeNos), JSON.stringify(relation.ufNos), `用户确认 · ${context}`, now, reviewer, now, "用户手动选择或填写后确认",
  );
  return now;
}

export function confirmUserMapping(input, targetInternal, context = "型号查询", reviewer = "当前用户", extras = {}) {
  const row = ensureUserModel(targetInternal);
  const relation = addConfirmedInput(row, input, context, reviewer, extras);
  const confirmedAt = recordUserConfirmation(input, row.id, relation, context, reviewer);
  if (relation.changed) insertVersion("型号库", "大厂型号匹配（自动更新）.xlsx", reviewer, {
    added: relation.createdModel ? 1 : 0,
    updated: relation.createdModel ? 0 : 1,
    relationsAdded: relation.addedOldModels.length + relation.addedOeNos.length + relation.addedUfNos.length,
    total: db.prepare("SELECT COUNT(*) AS count FROM models").get().count,
    context,
  });
  const updated = db.prepare("SELECT * FROM models WHERE id=?").get(row.id);
  return { model: serializeModel(updated, "用户确认"), confirmedAt, relation };
}

export function confirmMappingsBatch(items = [], context = "型号查询", reviewer = "当前用户") {
  const normalizedItems = (Array.isArray(items) ? items : []).map((item) => ({
    input: String(item?.input ?? item?.query ?? "").trim(),
    targetInternal: String(item?.targetInternal ?? item?.internal ?? "").trim(),
    oeNos: unique(Array.isArray(item?.oeNos) ? item.oeNos : []),
    ufNos: unique(Array.isArray(item?.ufNos) ? item.ufNos : []),
  })).filter((item) => item.input && item.targetInternal && item.targetInternal !== "未匹配" && item.targetInternal !== "待确认");
  if (!normalizedItems.length) throw new Error("补充表格中没有可确认的内部型号关系");
  const results = [];
  let added = 0;
  let updated = 0;
  let relationsAdded = 0;
  db.exec("BEGIN");
  try {
    for (const item of normalizedItems) {
      const row = ensureUserModel(item.targetInternal);
      const relation = addConfirmedInput(row, item.input, context, reviewer, item);
      recordUserConfirmation(item.input, row.id, relation, context, reviewer);
      if (relation.createdModel) added += 1; else if (relation.changed) updated += 1;
      relationsAdded += relation.addedOldModels.length + relation.addedOeNos.length + relation.addedUfNos.length;
      const updatedRow = db.prepare("SELECT * FROM models WHERE id=?").get(row.id);
      results.push({ query: item.input, targetInternal: item.targetInternal, model: serializeModel(updatedRow, "用户确认"), relation });
    }
    insertVersion("型号库", "搜索结果补充匹配.xlsx", reviewer, {
      added, updated, relationsAdded, total: db.prepare("SELECT COUNT(*) AS count FROM models").get().count, context,
    });
    db.exec("COMMIT");
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
  return { ok: true, count: results.length, results };
}

function resolveOeCandidates(models) {
  if (models.length <= 1 || !models.every((item) => item.matchedKind === "OE号")) return { candidates: models, selected: models.length === 1 ? models[0] : null, reason: "" };
  let candidates = models;
  let reason = "";
  const suffixPreferred = candidates.filter((item) => {
    const key = normalize(item.internal);
    return (key.endsWith("A1") || key.endsWith("B1")) && candidates.some((other) => normalize(other.internal) === key.slice(0, -1));
  });
  if (suffixPreferred.length) {
    candidates = suffixPreferred;
    const suffixes = unique(suffixPreferred.map((item) => normalize(item.internal).endsWith("A1") ? "A1" : "B1"));
    reason = `同基础型号优先选择 ${suffixes.join("/")} 结尾`;
    if (candidates.length === 1) return { candidates, selected: candidates[0], reason };
  }
  const short = candidates.filter((item) => normalize(item.internal).length <= 8);
  if (short.length && short.length < candidates.length) {
    candidates = short;
    reason = "优先选择 8 位及以内内部型号";
  }
  return { candidates, selected: candidates.length === 1 ? candidates[0] : null, reason };
}

function existingReview(query, modelId) {
  return db.prepare("SELECT * FROM reviews WHERE query=? AND COALESCE(model_id,0)=COALESCE(?,0) AND status='待校对' ORDER BY id DESC LIMIT 1").get(query, modelId || null);
}

async function createWebReview(query, local) {
  const existing = existingReview(query, local?.id);
  if (existing && String(existing.search_provider || "").startsWith(SEARCH_PROVIDER_PREFIX) && Date.now() - new Date(existing.searched_at).getTime() < 24 * 3600 * 1000) {
    const sources = json(existing.sources);
    let productImages = productImagesFromSources(sources);
    if (!productImages.length) {
      const imageSearch = await searchExternalModels(query, {
        vehicle: local?.vehicle || "",
        internal: local?.internal || "",
        oeNos: local?.oeNos || [],
        ufNos: local?.ufNos || [],
      });
      productImages = imageSearch.productImages || [];
    }
    return { id: existing.id, query: existing.query, oeNos: json(existing.proposed_oe), ufNos: json(existing.proposed_uf), verifiedOeNos: json(existing.verified_oe), verifiedUfNos: json(existing.verified_uf), platforms: json(existing.platforms), sources, productImages, confidence: existing.confidence, provider: existing.search_provider, searchedAt: existing.searched_at, status: existing.status, error: existing.note };
  }
  const web = await searchExternalModels(query, {
    vehicle: local?.vehicle || "",
    internal: local?.internal || "",
    oeNos: local?.oeNos || [],
    ufNos: local?.ufNos || [],
  });
  const oeNos = unique(web.oeNos.filter((value) => !local?.oeNos?.some((item) => normalize(item) === normalize(value))));
  const ufNos = unique(web.ufNos.filter((value) => !local?.ufNos?.some((item) => normalize(item) === normalize(value))));
  if (!web.sources.length && !oeNos.length && !ufNos.length) return { ...web, status: "搜索失败" };
  const verifiedOeNos = web.verifiedOeNos.filter((value) => oeNos.some((item) => normalize(item) === normalize(value)));
  const verifiedUfNos = web.verifiedUfNos.filter((value) => ufNos.some((item) => normalize(item) === normalize(value)));
  const provider = `${SEARCH_PROVIDER_PREFIX} · ${web.provider}`;
  db.prepare("UPDATE reviews SET status='已替换' WHERE query=? AND COALESCE(model_id,0)=COALESCE(?,0) AND status='待校对'").run(query, local?.id || null);
  const result = db.prepare("INSERT INTO reviews(query, model_id, proposed_oe, proposed_uf, verified_oe, verified_uf, platforms, sources, confidence, search_provider, searched_at, note) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)").run(
    query, local?.id || null, JSON.stringify(oeNos), JSON.stringify(ufNos), JSON.stringify(verifiedOeNos), JSON.stringify(verifiedUfNos), JSON.stringify(web.platforms), JSON.stringify(web.sources), web.confidence, provider, web.searchedAt, web.error || "",
  );
  return { id: Number(result.lastInsertRowid), ...web, provider, oeNos, ufNos, verifiedOeNos, verifiedUfNos, status: "待校对" };
}

function mergeWebSupplements(items = []) {
  const supplements = items.map((item) => item.webSupplement).filter(Boolean);
  if (!supplements.length) return null;
  const sources = [...new Map(supplements.flatMap((item) => item.sources || []).map((source) => [source.url || JSON.stringify(source), source])).values()];
  const platforms = unique(supplements.flatMap((item) => item.platforms || []));
  const oeNos = unique(supplements.flatMap((item) => item.oeNos || []));
  const ufNos = unique(supplements.flatMap((item) => item.ufNos || []));
  const verifiedOeNos = unique(supplements.flatMap((item) => item.verifiedOeNos || []));
  const verifiedUfNos = unique(supplements.flatMap((item) => item.verifiedUfNos || []));
  const productImages = productImagesFromSources(supplements.flatMap((item) => item.sources || []));
  const confidenceValues = supplements.map((item) => Number(item.confidence)).filter(Number.isFinite);
  return {
    ...supplements[0],
    query: items.map((item) => item.query).join(" / "),
    sources,
    platforms,
    oeNos,
    ufNos,
    verifiedOeNos,
    verifiedUfNos,
    productImages,
    confidence: confidenceValues.length ? Math.round(confidenceValues.reduce((sum, value) => sum + value, 0) / confidenceValues.length) : 0,
    status: supplements.some((item) => item.status === "待校对") ? "待校对" : supplements[0].status,
  };
}

function crossValidationFor(aliasResults = []) {
  const modelSets = aliasResults.map((item) => new Set((item.local || []).map((model) => model.id))).filter((set) => set.size);
  const commonModelIds = modelSets.length === aliasResults.length
    ? [...modelSets[0]].filter((id) => modelSets.every((set) => set.has(id)))
    : [];
  const commonInternal = commonModelIds.length
    ? aliasResults[0].local.find((model) => model.id === commonModelIds[0])?.internal || ""
    : "";
  return {
    items: aliasResults.map((item) => ({
      query: item.query,
      local: (item.local || []).map((model) => ({ id: model.id, internal: model.internal, matchedKind: model.matchedKind })),
      webSupplement: item.webSupplement || null,
    })),
    summary: { status: commonModelIds.length ? "通过" : "需校对", commonInternal, commonModelIds },
  };
}

async function searchModelsSingle(query, { includeWeb = true } = {}) {
  const lookup = findUniversalLocal(query);
  const found = lookup.models;
  const resolution = lookup.universal
    ? { candidates: found, selected: found.length === 1 ? found[0] : null, reason: found.length > 1 ? "通用型号分别匹配到不同内部型号，请选择" : found.length === 1 ? `通用型号 ${found[0].matchedAliases?.join(" / ")} 命中` : "" }
    : resolveOeCandidates(found);
  const local = resolution.selected ? [resolution.selected, ...found.filter((item) => item.id !== resolution.selected.id)] : found;
  const needsWeb = includeWeb && isValidModelInput(query) && (!local.length || local.some((item) => !item.oeNos.length || !item.ufNos.length));
  const webQuery = lookup.aliases[0] || query;
  const webContext = { vehicle: local[0]?.vehicle || "", internal: local[0]?.internal || "", oeNos: local[0]?.oeNos || [], ufNos: local[0]?.ufNos || [] };
  const webSupplement = needsWeb ? await createWebReview(webQuery, local[0]) : null;
  const imageSearch = includeWeb && local.length > 1 && !webSupplement ? await searchExternalModels(query, webContext) : null;
  return { query, aliases: lookup.aliases, universal: lookup.universal, local, preferredModelId: resolution.selected?.id || null, preferenceReason: resolution.reason, webSupplement, productImages: webSupplement?.productImages || imageSearch?.productImages || [], searchedAt: new Date().toISOString() };
}

export async function searchModels(query, { includeWeb = true } = {}) {
  const aliases = splitUniversalInputs(query);
  if (aliases.length < 2) return searchModelsSingle(query, { includeWeb });
  const aliasResults = await runPool(aliases, 8, (alias) => searchModelsSingle(alias, { includeWeb }));
  const lookup = findUniversalLocal(query);
  const collected = new Map();
  for (const result of aliasResults) {
    for (const model of result.local || []) {
      const existing = collected.get(model.id);
      if (existing) existing.matchedAliases = unique([...(existing.matchedAliases || []), result.query]);
      else collected.set(model.id, { ...model, matchedAliases: [result.query] });
    }
  }
  const found = [...collected.values()];
  const resolution = found.length === 1
    ? { selected: found[0], reason: `多个 OE 号交叉验证均命中 ${found[0].internal}` }
    : { selected: null, reason: found.length > 1 ? "多个 OE 号匹配到不同内部型号，请校对" : "" };
  const crossValidation = crossValidationFor(aliasResults);
  const webSupplement = mergeWebSupplements(aliasResults);
  let productImages = webSupplement?.productImages || [];
  if (includeWeb && found.length > 1 && !productImages.length) {
    const imageSearch = await searchExternalModels(query, { oeNos: lookup.aliases });
    productImages = imageSearch.productImages || [];
  }
  return {
    query,
    aliases: lookup.aliases,
    universal: true,
    local: resolution.selected ? [resolution.selected, ...found.filter((item) => item.id !== resolution.selected.id)] : found,
    preferredModelId: resolution.selected?.id || null,
    preferenceReason: resolution.reason,
    crossValidation: crossValidation.items,
    crossValidationSummary: crossValidation.summary,
    webSupplement,
    productImages,
    searchedAt: new Date().toISOString(),
  };
}

export async function searchModelsBatch(queries, { includeWeb = true } = {}) {
  const submitted = unique((queries || []).map((value) => String(value).trim()));
  const cleaned = submitted.filter(isValidModelInput).slice(0, MAX_BATCH_MODEL_QUERIES);
  const ignoredQueries = submitted.filter((value) => !isValidModelInput(value));
  if (!cleaned.length) throw new Error("请输入至少一个型号");
  const webEnabled = includeWeb && cleaned.length <= MAX_WEB_BATCH_QUERIES;
  const results = await runPool(cleaned, 10, (query) => searchModels(query, { includeWeb: webEnabled }));
  return {
    queries: cleaned,
    ignoredQueries,
    results,
    searchedAt: new Date().toISOString(),
    searchMode: webEnabled ? "local-and-web" : "local-only",
    searchModeNote: webEnabled ? "已完成本地匹配，并联网补齐缺失字段" : `本次共 ${cleaned.length} 个型号，已完成本地知识库批量匹配；为避免批量联网超时，缺失项未逐条联网搜索。`,
  };
}

export function cleanupInvalidPendingReviews(reviewer = "系统自动清理") {
  const invalid = db.prepare("SELECT id, query FROM reviews WHERE status='待校对' ORDER BY id").all().filter((row) => !isValidModelInput(row.query));
  if (!invalid.length) return { count: 0, ids: [] };
  const now = new Date().toISOString();
  const update = db.prepare("UPDATE reviews SET status='已驳回', reviewer=?, reviewed_at=?, note=? WHERE id=? AND status='待校对'");
  db.exec("BEGIN");
  try {
    invalid.forEach((row) => update.run(reviewer, now, "非型号内容自动过滤：数量、表头或年份", row.id));
    db.exec("COMMIT");
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
  return { count: invalid.length, ids: invalid.map((row) => row.id) };
}

export function getOverview() {
  const versions = db.prepare("SELECT * FROM versions ORDER BY id DESC LIMIT 6").all().map((row) => ({ ...row, summary: json(row.summary, {}) }));
  return {
    modelCount: db.prepare("SELECT COUNT(*) AS count FROM models").get().count,
    priceCount: db.prepare("SELECT COUNT(*) AS count FROM prices").get().count,
    catalogCount: db.prepare("SELECT COUNT(*) AS count FROM catalog_products").get().count,
    reviewCount: db.prepare("SELECT COUNT(*) AS count FROM reviews WHERE status='待校对'").get().count,
    quoteCount: db.prepare("SELECT COUNT(*) AS count FROM quote_jobs").get().count,
    modelVersion: versions.find((row) => row.type === "型号库")?.version || "M001",
    priceVersion: versions.find((row) => row.type === "价格库")?.version || "P001",
    catalogVersion: versions.find((row) => row.type === "产品目录")?.version || "C001",
    recentVersions: versions,
  };
}

export function listReviews() {
  const rows = [
    ...db.prepare("SELECT r.*, m.internal, m.vehicle FROM reviews r LEFT JOIN models m ON m.id=r.model_id WHERE r.status='待校对' ORDER BY r.id DESC LIMIT 100").all(),
    ...db.prepare("SELECT r.*, m.internal, m.vehicle FROM reviews r LEFT JOIN models m ON m.id=r.model_id WHERE r.status='已确认' ORDER BY r.reviewed_at DESC, r.id DESC LIMIT 100").all(),
  ];
  return rows.map((row) => {
    const verifiedUf = json(row.verified_uf);
    const rawProposedUf = json(row.proposed_uf);
    const noisyLegacyUf = /Google UF 专项校对 v4/.test(row.search_provider || "") && rawProposedUf.length > 2 && !verifiedUf.length;
    return {
      id: row.id, query: row.query, internal: row.internal || "待识别", vehicle: row.vehicle || "",
      proposedOe: json(row.proposed_oe).slice(0, 2), proposedUf: noisyLegacyUf ? [] : rawProposedUf.slice(0, 2), sources: json(row.sources).map(compactReviewSource), confidence: noisyLegacyUf ? 0 : row.confidence,
      verifiedOe: json(row.verified_oe), verifiedUf, platforms: json(row.platforms),
      provider: row.search_provider, searchedAt: row.searched_at, status: row.status, reviewer: row.reviewer, reviewedAt: row.reviewed_at, note: noisyLegacyUf ? "旧版搜索页推荐商品已过滤，请重新查询生成候选" : row.note,
    };
  });
}

export function reviewCandidate(id, action, reviewer = "当前用户", targetInternal = "", selectedOeNos = [], selectedUfNos = []) {
  const review = db.prepare("SELECT * FROM reviews WHERE id=?").get(id);
  if (!review) throw new Error("校对项不存在");
  if (review.status !== "待校对") return { ok: true, status: review.status };
  const status = action === "approve" ? "已确认" : "已驳回";
  let modelId = review.model_id;
  let createdModel = false;
  if (action === "approve" && targetInternal) {
    const ensured = ensureUserModel(targetInternal);
    modelId = ensured.id;
    createdModel = Boolean(ensured.__created);
  }
  if (action === "approve" && !modelId) throw new Error("请先填写要关联的内部型号");
  let confirmedMapping = null;
  if (action === "approve" && modelId) {
    const row = db.prepare("SELECT * FROM models WHERE id=?").get(modelId);
    const allowedOe = new Set(json(review.proposed_oe).map(normalize));
    const allowedUf = new Set(json(review.proposed_uf).map(normalize));
    const confirmedOe = unique((Array.isArray(selectedOeNos) ? selectedOeNos : []).filter((value) => allowedOe.has(normalize(value))));
    const confirmedUf = unique((Array.isArray(selectedUfNos) ? selectedUfNos : []).filter((value) => allowedUf.has(normalize(value))));
    const relation = addConfirmedInput({ ...row, __created: createdModel }, review.query, "二次校对", reviewer, {
      oeNos: confirmedOe, ufNos: confirmedUf, webOeNos: confirmedOe, webUfNos: confirmedUf,
    });
    db.prepare("UPDATE reviews SET proposed_oe=?, proposed_uf=?, verified_oe=?, verified_uf=? WHERE id=?").run(JSON.stringify(relation.oeNos), JSON.stringify(relation.ufNos), JSON.stringify(relation.oeNos), JSON.stringify(relation.ufNos), id);
    if (relation.changed) insertVersion("型号库", "大厂型号匹配（自动更新）.xlsx", reviewer, {
      added: relation.createdModel ? 1 : 0, updated: relation.createdModel ? 0 : 1,
      relationsAdded: relation.addedOldModels.length + relation.addedOeNos.length + relation.addedUfNos.length,
      total: db.prepare("SELECT COUNT(*) AS count FROM models").get().count,
      context: "二次校对",
    });
    confirmedMapping = { internal: row.internal, oeNos: relation.oeNos, ufNos: relation.ufNos };
  }
  const now = new Date().toISOString();
  db.prepare("UPDATE reviews SET status=?, reviewer=?, reviewed_at=?, model_id=COALESCE(?,model_id) WHERE id=?").run(status, reviewer, now, modelId || null, id);
  return { ok: true, status, reviewer, reviewedAt: now, confirmedMapping };
}

function detectInquirySheet(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: false });
  let headerRow = -1;
  let modelColumn = -1;
  let bestScore = 0;
  const normalizedHeader = (value) => String(value || "").trim().toUpperCase().replace(/[\s._/\\-]+/g, "");
  const isExplicitOeHeader = (value) => /^(?:OENUMBER|OENO|OEMNUMBER|OEMNO|OECODE|OEMCODE|OE号|OEM号)$/.test(normalizedHeader(value));
  const headerScore = (value) => {
    const key = normalizedHeader(value);
    if (!key || /^(?:QTY|QUANTITY|PCS|数量|单价|价格|TOTALPRICE|AMOUNT|NO|序号)$/.test(key)) return 0;
    // OE / UF 是报价最可靠的查询键，必须高于通用的 Item No / 序号列。
    if (/^(?:OENUMBER|OENO|OEMNUMBER|OEMNO|OE号|OEM号)$/.test(key)) return 100;
    if (/^(?:UFNUMBER|UFNO|UF号)$/.test(key)) return 96;
    if (/^(?:内部型号|对外型号|产品型号|客户型号|货品编码|产品编码|零件号|配件号)$/.test(key)) return 90;
    if (/^(?:PARTNUMBER|PARTNO|MODELNUMBER|MODELNO|OE|OEM|UF|MODEL|型号|货号)$/.test(key)) return 82;
    if (/(?:OE号|OEM号|UF号|内部型号|对外型号|产品型号|客户型号|PARTNO|PARTNUMBER|MODELNO|MODELNUMBER|货品编码|产品编码|零件号|配件号)/.test(key)) return 70;
    if (/^(?:ITEMNUMBER|ITEMNO)$/.test(key)) return 10;
    return 0;
  };
  for (let row = 0; row < Math.min(15, rows.length); row += 1) {
    for (let column = 0; column < rows[row].length; column += 1) {
      const baseScore = headerScore(rows[row][column]);
      const samples = rows.slice(row + 1, row + 31).map((item) => item[column]).filter((cell) => String(cell || "").trim());
      const validCount = samples.filter(isValidModelInput).length;
      // 客户表中的 OEM No 就是 OE 号；一旦存在有效数据，强制高于 Item No、序号等其他列。
      const score = validCount ? baseScore + (isExplicitOeHeader(rows[row][column]) ? 1000 : 0) + Math.round((validCount / Math.max(1, samples.length)) * 10) : 0;
      if (score > bestScore) { bestScore = score; headerRow = row; modelColumn = column; }
    }
  }
  if (modelColumn < 0) {
    const width = Math.max(0, ...rows.slice(0, 30).map((row) => row.length));
    let bestCount = 0;
    for (let column = 0; column < width; column += 1) {
      const samples = rows.slice(0, 30).map((row) => row[column]).filter((cell) => String(cell || "").trim());
      const count = samples.filter(isValidModelInput).length;
      if (count >= 2 && count / samples.length >= 0.6 && count > bestCount) {
        bestCount = count;
        headerRow = -1;
        modelColumn = column;
      }
    }
  }
  if (modelColumn < 0) throw new Error("未识别到型号 / OE 列，请确认询价表包含 OE NUMBER、内部型号或产品型号表头");
  const headers = (rows[headerRow] || []).map((cell, index) => String(cell).trim() || `列${index + 1}`);
  if (headerRow < 0) {
    const width = Math.max(0, ...rows.map((row) => row.length));
    return { rows, headerRow, modelColumn, headers: Array.from({ length: width }, (_, index) => `列${index + 1}`), sheetName };
  }
  return { rows, headerRow, modelColumn, headers, sheetName };
}

function normalizePricings(value) {
  const source = Array.isArray(value) ? value : [value || {}];
  const seen = new Set();
  const result = [];
  for (const item of source) {
    const pricing = {
      packaging: ["彩盒", "中性"].includes(item?.packaging) ? item.packaging : "彩盒",
      tax: ["含税", "未税"].includes(item?.tax) ? item.tax : "未税",
      tier: ["NP", "LP", "SLP", "2S", "3S", "4S"].includes(item?.tier) ? item.tier : "4S",
    };
    const key = `${pricing.tier}|${pricing.packaging}|${pricing.tax}`;
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(pricing);
  }
  return result.length ? result.slice(0, 12) : [{ packaging: "彩盒", tax: "未税", tier: "4S" }];
}

function pricingLabel(pricing) {
  return `${pricing.tier} ${pricing.packaging}${pricing.tax}`;
}

function pricesForInternal(internal, pricings) {
  return pricings.map((pricing) => {
    const price = db.prepare("SELECT * FROM prices WHERE normalized_internal=? AND packaging=? AND tax=? AND tier=?").get(normalize(internal), pricing.packaging, pricing.tax, pricing.tier);
    return {
      ...pricing,
      label: pricingLabel(pricing),
      price: price?.price ?? null,
      attribute: price?.attribute || "",
      source: price ? `${price.source_sheet} · ${price.source_column}` : "",
    };
  });
}

function emptyPrices(pricings) {
  return pricings.map((pricing) => ({ ...pricing, label: pricingLabel(pricing), price: null, attribute: "", source: "" }));
}

function attributeFromPrices(prices = []) {
  return prices.find((item) => item.attribute)?.attribute || "";
}

function missingPriceIssue(prices, prefix = "已匹配内部型号") {
  const missing = prices.filter((item) => item.price === null).map((item) => item.label);
  return missing.length ? `${prefix}，但以下价格类型无价格：${missing.join("、")}` : "";
}

async function mapQuoteLine(input, original, pricings) {
  const lookup = findUniversalLocal(input, { exact: true });
  const found = lookup.models;
  const crossValidation = lookup.universal ? crossValidationFor(lookup.aliases.map((query) => ({ query, local: findLocalExact(query), webSupplement: null }))) : null;
  const validationFields = crossValidation ? { crossValidation: crossValidation.items, crossValidationSummary: crossValidation.summary } : {};
  const resolution = lookup.universal
    ? { candidates: found, selected: found.length === 1 ? found[0] : null, reason: found.length > 1 ? "通用型号分别匹配到不同内部型号，请选择" : found.length === 1 ? `通用型号 ${found[0].matchedAliases?.join(" / ")} 命中` : "" }
    : resolveOeCandidates(found);
  const local = resolution.candidates;
  const candidateOptions = local.map((model) => {
    const catalog = catalogForInternal(model.internal);
    const displayInternal = normalizeRedShellInternal(model.internal, catalog);
    const prices = pricesForInternal(displayInternal, pricings);
    return {
      modelId: model.id, internal: displayInternal, matchedOriginalInternal: displayInternal !== model.internal ? model.internal : "", redShellAdjusted: displayInternal !== model.internal, oe: model.oeNos.join(" / "), uf: model.ufNos.join(" / "), vehicle: model.vehicle, vehicleEnglish: vehicleToEnglish(model.vehicle),
      matchedKind: model.matchedKind, prices, price: prices[0]?.price ?? null, priceSource: prices[0]?.source || "",
      attribute: attributeFromPrices(prices),
      catalogImage: catalog?.image || "", catalogName: catalog?.commonName || "", catalogSource: catalog?.sourceSheet || "",
      modelSource: model.sourceType === "original" ? "原始匹配表" : model.sourceType === "user_confirmed" ? "用户确认" : "外网补充",
    };
  });
  const multipleImageSearch = candidateOptions.length > 1 && isValidModelInput(input)
    ? await searchExternalModels(input, { oeNos: candidateOptions.flatMap((item) => item.oe.split(" / ")).slice(0, 6) })
    : null;
  if (resolution.selected && candidateOptions.length === 1) {
    const selected = candidateOptions[0];
    return {
      original, input, internal: selected.internal, oe: selected.oe, uf: selected.uf, vehicle: selected.vehicle, vehicleEnglish: selected.vehicleEnglish, attribute: selected.attribute, prices: selected.prices, price: selected.price,
      catalogImage: selected.catalogImage, catalogName: selected.catalogName, catalogSource: selected.catalogSource,
      priceSource: selected.priceSource, modelSource: selected.modelSource, matchedOriginalInternal: selected.matchedOriginalInternal || "", redShellAdjusted: Boolean(selected.redShellAdjusted), reviewStatus: "已确认", selectedModelId: selected.modelId,
      matchedKind: selected.matchedKind, candidateOptions, selectionReason: resolution.reason,
      ...validationFields,
      issue: missingPriceIssue(selected.prices),
    };
  }
  if (candidateOptions.length > 1) {
    return {
      original, input, internal: "待选择", oe: "", uf: "", vehicle: "", vehicleEnglish: "", attribute: "", prices: emptyPrices(pricings), price: null, priceSource: "", catalogImage: "", catalogName: "", catalogSource: "", modelSource: "原始匹配表",
      reviewStatus: "待选择", issue: "匹配到多个内部型号，请选择", matchedKind: candidateOptions[0].matchedKind, candidateOptions, selectionReason: resolution.reason, productImages: multipleImageSearch?.productImages || [], ...validationFields,
    };
  }
  if (candidateOptions.length === 1) {
    const selected = candidateOptions[0];
    return {
      original, input, internal: selected.internal, oe: selected.oe, uf: selected.uf, vehicle: selected.vehicle, vehicleEnglish: selected.vehicleEnglish, attribute: selected.attribute, prices: selected.prices, price: selected.price,
      catalogImage: selected.catalogImage, catalogName: selected.catalogName, catalogSource: selected.catalogSource,
      priceSource: selected.priceSource, modelSource: selected.modelSource, matchedOriginalInternal: selected.matchedOriginalInternal || "", redShellAdjusted: Boolean(selected.redShellAdjusted), reviewStatus: "已确认", selectedModelId: selected.modelId,
      matchedKind: selected.matchedKind, candidateOptions, ...validationFields, issue: missingPriceIssue(selected.prices),
    };
  }
  if (!isValidModelInput(input)) {
    return { original, input, internal: "", oe: "", uf: "", vehicle: "", vehicleEnglish: "", attribute: "", prices: emptyPrices(pricings), price: null, priceSource: "", catalogImage: "", catalogName: "", catalogSource: "", modelSource: "未匹配", reviewStatus: "", issue: "该内容不像有效型号，未发起联网搜索", candidateOptions: [], productImages: [] };
  }
  const webSupplement = await createWebReview(input, null);
  const issue = webSupplement?.sources?.length ? "本地型号库未匹配，已发起联网补充" : "本地与外网均未找到可靠型号";
  return {
    original, input, internal: "待确认", oe: webSupplement?.oeNos?.join(" / ") || "", uf: webSupplement?.ufNos?.join(" / ") || "", vehicle: "", vehicleEnglish: "", attribute: "", prices: emptyPrices(pricings), price: null,
    catalogImage: "", catalogName: "", catalogSource: "",
    priceSource: "", modelSource: webSupplement?.sources?.length ? "外网补充" : "未匹配", reviewStatus: webSupplement?.status || "待补充", issue, candidateOptions: [], productImages: webSupplement?.productImages || [], ...validationFields,
  };
}

function refreshQuoteMetrics(result) {
  result.total = result.lines.length;
  result.matched = result.lines.filter((line) => line.prices?.length ? line.prices.every((item) => item.price !== null) : line.price !== null).length;
  result.needsSelection = result.lines.filter((line) => line.reviewStatus === "待选择").length;
  result.needsReview = result.lines.filter((line) => line.reviewStatus === "待校对").length;
  result.unmatched = result.lines.filter((line) => {
    const missing = line.prices?.length ? line.prices.some((item) => item.price === null) : line.price === null;
    return missing && !["待选择", "待校对"].includes(line.reviewStatus);
  }).length;
  return result;
}

async function runPool(items, limit, worker) {
  const results = new Array(items.length);
  let index = 0;
  async function runner() {
    while (index < items.length) {
      const current = index;
      index += 1;
      results[current] = await worker(items[current], current);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, runner));
  return results;
}

export async function previewQuote(buffer, filename, pricingInput) {
  const pricings = normalizePricings(pricingInput);
  const inquiry = detectInquirySheet(buffer);
  const sourceRows = inquiry.rows.slice(inquiry.headerRow + 1).map((row, index) => ({ row, sourceRow: inquiry.headerRow + 1 + index })).filter((item) => item.row.some((cell) => String(cell).trim()));
  const lines = await runPool(sourceRows.slice(0, 500), 4, async ({ row, sourceRow }) => {
    const original = Object.fromEntries(inquiry.headers.map((header, index) => [header, row[index] ?? ""]));
    const input = String(row[inquiry.modelColumn] || "").trim();
    const line = input ? await mapQuoteLine(input, original, pricings) : { original, input: "", internal: "", oe: "", uf: "", vehicle: "", vehicleEnglish: "", attribute: "", prices: emptyPrices(pricings), price: null, priceSource: "", catalogImage: "", catalogName: "", catalogSource: "", modelSource: "未匹配", reviewStatus: "", issue: "型号为空" };
    return { ...line, sourceRow };
  });
  const id = randomUUID();
  const pricingLabels = pricings.map(pricingLabel);
  const pricingLabelText = pricingLabels.join("、");
  const result = {
    id, filename, sheetName: inquiry.sheetName, detectedColumn: inquiry.headers[inquiry.modelColumn], pricingLabel: pricingLabelText,
    pricingLabels, pricings, pricing: pricings[0],
    lines,
  };
  refreshQuoteMetrics(result);
  db.prepare("INSERT INTO quote_jobs(id, filename, pricing_label, result_json, created_at) VALUES (?, ?, ?, ?, ?)").run(id, filename, pricingLabelText, JSON.stringify(result), new Date().toISOString());
  return result;
}

export function selectQuoteCandidate(id, lineIndex, modelId) {
  const row = db.prepare("SELECT * FROM quote_jobs WHERE id=?").get(id);
  if (!row) throw new Error("报价任务不存在");
  const result = json(row.result_json, {});
  const line = result.lines?.[Number(lineIndex)];
  if (!line) throw new Error("报价行不存在");
  const selected = line.candidateOptions?.find((item) => Number(item.modelId) === Number(modelId));
  if (!selected) throw new Error("所选内部型号不在候选列表中");
  Object.assign(line, {
    internal: selected.internal, oe: selected.oe, uf: selected.uf, vehicle: selected.vehicle, vehicleEnglish: selected.vehicleEnglish, attribute: selected.attribute || attributeFromPrices(selected.prices), prices: selected.prices || [], price: selected.price,
    catalogImage: selected.catalogImage || "", catalogName: selected.catalogName || "", catalogSource: selected.catalogSource || "",
    priceSource: selected.priceSource, modelSource: selected.modelSource, selectedModelId: selected.modelId, matchedOriginalInternal: selected.matchedOriginalInternal || "", redShellAdjusted: Boolean(selected.redShellAdjusted),
    matchedKind: selected.matchedKind, reviewStatus: "已确认", issue: missingPriceIssue(selected.prices || [], "所选内部型号"),
  });
  confirmUserMapping(line.input, selected.internal, "报价", "当前用户");
  refreshQuoteMetrics(result);
  db.prepare("UPDATE quote_jobs SET result_json=? WHERE id=?").run(JSON.stringify(result), id);
  return result;
}

export function confirmQuoteLine(id, lineIndex, targetInternal, reviewer = "当前用户") {
  const row = db.prepare("SELECT * FROM quote_jobs WHERE id=?").get(id);
  if (!row) throw new Error("报价任务不存在");
  const result = json(row.result_json, {});
  const line = result.lines?.[Number(lineIndex)];
  if (!line) throw new Error("报价行不存在");
  const confirmation = confirmUserMapping(line.input, targetInternal, "报价", reviewer);
  const pricings = normalizePricings(result.pricings || result.pricing || {});
  applyConfirmedQuoteModel(line, confirmation.model, pricings, "用户确认");
  line.selectionReason = "用户手动修改并确认";
  refreshQuoteMetrics(result);
  db.prepare("UPDATE quote_jobs SET result_json=? WHERE id=?").run(JSON.stringify(result), id);
  return result;
}

function applyConfirmedQuoteModel(line, model, pricings, matchedKind = "补充导入") {
  const catalog = catalogForInternal(model.internal);
  const displayInternal = normalizeRedShellInternal(model.internal, catalog);
  const prices = pricesForInternal(displayInternal, pricings);
  Object.assign(line, {
    internal: displayInternal,
    matchedOriginalInternal: displayInternal !== model.internal ? model.internal : "",
    redShellAdjusted: displayInternal !== model.internal,
    oe: model.oeNos.join(" / "),
    uf: model.ufNos.join(" / "),
    vehicle: model.vehicle,
    vehicleEnglish: model.vehicleEnglish || vehicleToEnglish(model.vehicle),
    attribute: attributeFromPrices(prices),
    prices,
    price: prices[0]?.price ?? null,
    priceSource: prices[0]?.source || "",
    catalogImage: catalog?.image || "",
    catalogName: catalog?.commonName || "",
    catalogSource: catalog?.sourceSheet || "",
    modelSource: "用户确认",
    selectedModelId: model.id,
    matchedKind,
    reviewStatus: "已确认",
    candidateOptions: [],
    selectionReason: "上传补充表确认",
    supplemented: true,
    issue: missingPriceIssue(prices, "已确认内部型号"),
  });
  return line;
}

export function supplementQuoteJob(id, items = [], reviewer = "当前用户") {
  const row = db.prepare("SELECT * FROM quote_jobs WHERE id=?").get(id);
  if (!row) throw new Error("报价任务不存在");
  const result = json(row.result_json, {});
  const pricings = normalizePricings(result.pricings || result.pricing || {});
  const updates = Array.isArray(items) ? items : [];
  let count = 0;
  for (const item of updates) {
    const lineIndex = Number(item?.lineIndex);
    const line = Number.isInteger(lineIndex) ? result.lines?.[lineIndex] : null;
    if (!line) continue;
    const targetInternal = String(item?.targetInternal ?? item?.internal ?? "").trim();
    if (!targetInternal || targetInternal === "未匹配" || targetInternal === "待确认") continue;
    const input = String(item?.input || line.input || "").trim();
    if (!input) continue;
    const confirmation = confirmUserMapping(input, targetInternal, "报价补充", reviewer, {
      oeNos: unique(Array.isArray(item?.oeNos) ? item.oeNos : []),
      ufNos: unique(Array.isArray(item?.ufNos) ? item.ufNos : []),
    });
    applyConfirmedQuoteModel(line, confirmation.model, pricings);
    count += 1;
  }
  if (!count) throw new Error("没有识别到可合并的报价补充关系");
  refreshQuoteMetrics(result);
  db.prepare("UPDATE quote_jobs SET result_json=? WHERE id=?").run(JSON.stringify(result), id);
  return result;
}

export function attachQuoteImages(id, images = []) {
  const row = db.prepare("SELECT * FROM quote_jobs WHERE id=?").get(id);
  if (!row) throw new Error("报价任务不存在");
  const result = json(row.result_json, {});
  for (const item of images.slice(0, 200)) {
    const line = result.lines?.[Number(item.lineIndex)];
    if (!line) continue;
    const image = String(item.image || "");
    line.customerImage = /^data:image\/(?:jpeg|jpg|png);base64,/i.test(image) ? image : "";
    line.customerImageValidated = Boolean(line.customerImage && item.validated === true);
    line.customerImageStatus = String(item.status || "").slice(0, 30);
  }
  db.prepare("UPDATE quote_jobs SET result_json=? WHERE id=?").run(JSON.stringify(result), id);
  return result;
}

function dataUrlParts(value = "") {
  const match = String(value).match(/^data:image\/(png|jpe?g);base64,(.+)$/i);
  return match ? { extension: match[1].toLowerCase() === "png" ? "png" : "jpeg", base64: value } : null;
}

function localImageParts(value = "") {
  const data = dataUrlParts(value);
  if (data) return data;
  if (!String(value).startsWith("/")) return null;
  const filePath = path.join(publicDir, String(value).replace(/^\/+/, ""));
  if (!fs.existsSync(filePath)) return null;
  const extension = path.extname(filePath).toLowerCase() === ".png" ? "png" : "jpeg";
  return { extension, buffer: fs.readFileSync(filePath) };
}

function addExcelImage(workbook, worksheet, value, columnIndex, rowIndex) {
  const image = localImageParts(value);
  if (!image) return false;
  const imageId = workbook.addImage(image.base64 ? { base64: image.base64, extension: image.extension } : { buffer: image.buffer, extension: image.extension });
  worksheet.addImage(imageId, { tl: { col: columnIndex - 1 + .08, row: rowIndex - 1 + .08 }, ext: { width: 96, height: 66 } });
  return true;
}

export async function exportQuote(id) {
  const row = db.prepare("SELECT * FROM quote_jobs WHERE id=?").get(id);
  if (!row) throw new Error("报价任务不存在");
  const result = json(row.result_json, {});
  const outputRows = result.lines.map((line) => {
    const dynamicPrices = {};
    (line.prices || [{ label: result.pricingLabel, price: line.price, source: line.priceSource }]).forEach((item, index) => {
      const label = item.label || result.pricingLabels?.[index] || result.pricingLabel || `价格类型${index + 1}`;
      dynamicPrices[`${label}报价`] = item.price ?? "";
      dynamicPrices[`${label}价格来源`] = item.source || "";
    });
    return {
      ...line.original, "客户询价图片": line.customerImageValidated && line.customerImage ? "已嵌入" : "", "匹配内部型号": line.internal, "红色外壳原型号": line.matchedOriginalInternal || "", "属性（自制/调货）": line.attribute || "", "产品目录图片": line.catalogImage ? "已嵌入" : "暂无图片", "匹配OE号": line.oe, "匹配UF号": line.uf, "适用车型": line.vehicle, "适用车型（英文）": line.vehicleEnglish || vehicleToEnglish(line.vehicle),
      "匹配依据": line.matchedKind ? `通过${line.matchedKind}命中` : "", ...dynamicPrices, "型号来源": line.modelSource, "校对状态": line.reviewStatus, "异常说明": line.issue,
    };
  });
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("报价结果", { views: [{ state: "frozen", ySplit: 1 }] });
  const keys = Object.keys(outputRows[0] || { 型号: "" });
  worksheet.columns = keys.map((key) => ({ header: key, key, width: key.includes("图片") ? 17 : Math.min(36, Math.max(13, key.length * 2 + 4)) }));
  worksheet.getRow(1).height = 26;
  worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  worksheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF145DF5" } };
  outputRows.forEach((item, index) => {
    const excelRow = index + 2;
    worksheet.addRow(item);
    worksheet.getRow(excelRow).height = 56;
    worksheet.getRow(excelRow).alignment = { vertical: "middle", wrapText: true };
    const line = result.lines[index];
    if (line.customerImageValidated) addExcelImage(workbook, worksheet, line.customerImage, keys.indexOf("客户询价图片") + 1, excelRow);
    addExcelImage(workbook, worksheet, line.catalogImage, keys.indexOf("产品目录图片") + 1, excelRow);
  });
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

const pendingImports = new Map();

export function previewKnowledgeImport(buffer, filename, type) {
  const token = randomUUID();
  if (type === "relations") {
    const parsed = parseMappingReview(buffer);
    const prepared = prepareMappingReview(parsed, filename);
    pendingImports.set(token, { buffer, filename, type, parsed, summary: prepared.summary, createdAt: Date.now() });
    return { token, type, filename, ...prepared.summary };
  }
  const parsed = type === "models" ? parseModels(buffer) : type === "catalog" ? parseCatalog(buffer) : parsePrices(buffer);
  const existingCount = type === "models" ? db.prepare("SELECT COUNT(*) AS count FROM models").get().count : type === "catalog" ? db.prepare("SELECT COUNT(*) AS count FROM catalog_products").get().count : db.prepare("SELECT COUNT(*) AS count FROM prices").get().count;
  const incomingKeys = new Set(parsed.map((item) => type === "models" ? normalize(item.internal) : type === "catalog" ? item.normalizedInternal : `${normalize(item.internal)}|${item.packaging}|${item.tax}|${item.tier}`));
  const existingRows = type === "models" ? db.prepare("SELECT * FROM models").all() : type === "catalog" ? db.prepare("SELECT * FROM catalog_products").all() : db.prepare("SELECT *, normalized_internal || '|' || packaging || '|' || tax || '|' || tier AS key FROM prices").all();
  const existingKeys = new Set(existingRows.map((row) => type === "models" ? row.normalized_internal : type === "catalog" ? row.normalized_internal : row.key));
  const added = [...incomingKeys].filter((key) => !existingKeys.has(key)).length;
  const incomingByKey = new Map(parsed.map((item) => [type === "models" ? normalize(item.internal) : type === "catalog" ? item.normalizedInternal : `${normalize(item.internal)}|${item.packaging}|${item.tax}|${item.tier}`, item]));
  const existingByKey = new Map();
  for (const row of existingRows) {
    const key = type === "models" ? row.normalized_internal : type === "catalog" ? row.normalized_internal : row.key;
    if (type === "prices") {
      const rows = existingByKey.get(key) || [];
      rows.push(row);
      existingByKey.set(key, rows);
    } else {
      existingByKey.set(key, row);
    }
  }
  let supplemented = 0;
  let preserved = 0;
  for (const [key, item] of incomingByKey) {
    const previous = existingByKey.get(key);
    if (!previous) continue;
    if (type === "models") {
      const hasNew = [
        [json(previous.product_codes), item.productCodes], [json(previous.old_models), item.oldModels],
        [json(previous.oe_nos), item.oeNos], [json(previous.uf_nos), item.ufNos],
      ].some(([before, after]) => after.some((value) => !before.includes(value)))
        || mergeTextValues(previous.vehicle, item.vehicle) !== String(previous.vehicle || "")
        || mergeTextValues(previous.category, item.category) !== String(previous.category || "");
      if (hasNew) supplemented += 1;
      else preserved += 1;
    } else if (type === "catalog") {
      const changed = [
        [previous.internal, mergeTextValues(previous.internal, item.internal)], [previous.common_name, mergeTextValues(previous.common_name, item.commonName)],
        [previous.vehicle, mergeTextValues(previous.vehicle, item.vehicle)], [previous.oe, mergeTextValues(previous.oe, item.oe)],
        [previous.note, mergeTextValues(previous.note, item.note)], [previous.source_sheet, mergeTextValues(previous.source_sheet, item.sourceSheet)],
        [previous.source_row, previous.source_row || item.sourceRow], [previous.source_label, mergeTextValues(previous.source_label, filename)],
      ].some(([before, after]) => String(before || "") !== String(after || ""));
      if (changed) supplemented += 1;
      else preserved += 1;
    } else {
      const rows = previous || [];
      const sameValue = rows.some((row) => Number(row.price) === Number(item.price) && String(row.attribute || "") === String(item.attribute || "") && String(row.source_sheet || "") === String(item.sourceSheet || "") && String(row.source_column || "") === String(item.sourceColumn || ""));
      if (sameValue) preserved += 1;
      else supplemented += 1;
    }
  }
  const summary = { incoming: parsed.length, existing: existingCount, added, supplemented, preserved, conflicts: 0 };
  pendingImports.set(token, { buffer, filename, type, parsed, summary, createdAt: Date.now() });
  return { token, type, filename, ...summary };
}

export function applyKnowledgeImport(token, operator = "当前用户") {
  const pending = pendingImports.get(token);
  if (!pending) throw new Error("导入预览已失效，请重新上传");
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const safeName = pending.filename.replace(/[^\p{L}\p{N}._-]/gu, "_");
  fs.writeFileSync(path.join(uploadDir, `${stamp}-${safeName}`), pending.buffer);
  const relationResult = pending.type === "relations" ? applyMappingReviewParsed(pending.parsed, pending.filename, operator) : null;
  const summary = relationResult?.summary || (pending.type === "models" ? upsertModels(pending.parsed, pending.filename) : pending.type === "catalog" ? upsertCatalog(pending.parsed, pending.filename.replace(/\.catalog\.json$/i, "")) : upsertPrices(pending.parsed));
  const version = relationResult?.version || insertVersion(pending.type === "models" ? "型号库" : pending.type === "catalog" ? "产品目录" : "价格库", pending.filename.replace(/\.catalog\.json$/i, ""), operator, summary);
  pendingImports.delete(token);
  return { ok: true, version, summary };
}

export function applyKnowledgeFile(buffer, filename, type, operator = "当前用户") {
  if (type === "relations") {
    const parsed = parseMappingReview(buffer);
    if (!(parsed.decisions.length || parsed.mergedRelations.length)) throw new Error("文件中没有识别到关系核验数据");
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const safeName = filename.replace(/[^\p{L}\p{N}._-]/gu, "_");
    fs.writeFileSync(path.join(uploadDir, `${stamp}-${safeName}`), buffer);
    const result = applyMappingReviewParsed(parsed, filename, operator);
    return { ok: true, version: result.version, summary: result.summary };
  }
  const parsed = type === "models" ? parseModels(buffer) : type === "catalog" ? parseCatalog(buffer) : parsePrices(buffer);
  if (!parsed.length) throw new Error("文件中没有识别到可导入的数据");
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const safeName = filename.replace(/[^\p{L}\p{N}._-]/gu, "_");
  fs.writeFileSync(path.join(uploadDir, `${stamp}-${safeName}`), buffer);
  const summary = type === "models" ? upsertModels(parsed, filename) : type === "catalog" ? upsertCatalog(parsed, filename.replace(/\.catalog\.json$/i, "")) : upsertPrices(parsed);
  const version = insertVersion(type === "models" ? "型号库" : type === "catalog" ? "产品目录" : "价格库", filename.replace(/\.catalog\.json$/i, ""), operator, summary);
  return { ok: true, version, summary };
}

export function listVersions() {
  return db.prepare("SELECT * FROM versions ORDER BY id DESC LIMIT 30").all().map((row) => ({ ...row, summary: json(row.summary, {}) }));
}

export function syncConfirmedMappings() {
  const result = migrateConfirmedReviewMappings();
  return { ok: true, ...result };
}

export async function exportKnowledge(type) {
  if (type === "catalog") {
    const rows = db.prepare("SELECT * FROM catalog_products ORDER BY normalized_internal").all();
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("产品目录", { views: [{ state: "frozen", ySplit: 1 }] });
    worksheet.columns = [
      { header: "画册型号", key: "internal", width: 18 },
      { header: "产品图片", key: "image", width: 18 },
      { header: "通用名", key: "commonName", width: 24 },
      { header: "适用车型", key: "vehicle", width: 42 },
      { header: "OE NO.", key: "oe", width: 34 },
      { header: "备注", key: "note", width: 20 },
      { header: "来源工作表", key: "sourceSheet", width: 22 },
    ];
    worksheet.getRow(1).height = 26;
    worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    worksheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF145DF5" } };
    rows.forEach((row, index) => {
      const excelRow = index + 2;
      worksheet.addRow({ internal: row.internal, image: row.image_data || row.image_url ? "已嵌入" : "暂无图片", commonName: row.common_name, vehicle: row.vehicle, oe: row.oe, note: row.note, sourceSheet: row.source_sheet });
      worksheet.getRow(excelRow).height = 56;
      worksheet.getRow(excelRow).alignment = { vertical: "middle", wrapText: true };
      addExcelImage(workbook, worksheet, row.image_data || row.image_url, 2, excelRow);
    });
    return Buffer.from(await workbook.xlsx.writeBuffer());
  }
  if (type === "models") {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("大厂型号匹配", { views: [{ state: "frozen", ySplit: 1 }] });
    worksheet.columns = [
      { header: "货品编码", key: "productCodes", width: 20 },
      { header: "对外型号", key: "internal", width: 18 },
      { header: "老型号", key: "oldModels", width: 22 },
      { header: "OE号", key: "oeNos", width: 50 },
      { header: "UF号", key: "ufNos", width: 28 },
      { header: "适用车型", key: "vehicle", width: 34 },
      { header: "货品类别", key: "category", width: 18 },
      { header: "数据来源", key: "sourceLabel", width: 26 },
      { header: "外网确认OE", key: "webOeNos", width: 30 },
      { header: "外网确认UF", key: "webUfNos", width: 26 },
    ];
    worksheet.getRow(1).height = 26;
    worksheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    worksheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF145DF5" } };
    const confirmed = new Map();
    for (const item of db.prepare("SELECT model_id, kind, normalized_value FROM confirmed_relations WHERE is_new=1").all()) {
      const key = `${item.model_id}|${item.kind}`;
      if (!confirmed.has(key)) confirmed.set(key, new Set());
      confirmed.get(key).add(item.normalized_value);
    }
    const richList = (values, marked = new Set()) => {
      const list = unique(values);
      if (!list.length) return "";
      return { richText: list.flatMap((value, index) => [
        ...(index ? [{ text: " / ", font: { color: { argb: "FF344563" } } }] : []),
        { text: String(value), font: { color: { argb: marked.has(normalize(value)) ? "FFFFC000" : "FF344563" }, bold: marked.has(normalize(value)) } },
      ]) };
    };
    for (const row of db.prepare("SELECT * FROM models ORDER BY normalized_internal").all()) {
      const excelRow = worksheet.addRow({
        productCodes: json(row.product_codes).join(" / "), internal: row.internal, oldModels: "", oeNos: "", ufNos: "",
        vehicle: row.vehicle, category: row.category, sourceLabel: row.source_label,
        webOeNos: json(row.web_oe_nos).join(" / "), webUfNos: json(row.web_uf_nos).join(" / "),
      });
      const internalMarked = confirmed.get(`${row.id}|内部型号`)?.has(normalize(row.internal));
      if (internalMarked) excelRow.getCell("internal").font = { color: { argb: "FFFFC000" }, bold: true };
      excelRow.getCell("oldModels").value = richList(json(row.old_models), confirmed.get(`${row.id}|老型号`) || new Set());
      excelRow.getCell("oeNos").value = richList(json(row.oe_nos), confirmed.get(`${row.id}|OE号`) || new Set());
      excelRow.getCell("ufNos").value = richList(json(row.uf_nos), confirmed.get(`${row.id}|UF号`) || new Set());
      excelRow.alignment = { vertical: "middle", wrapText: true };
    }
    return Buffer.from(await workbook.xlsx.writeBuffer());
  } else {
    const workbook = XLSX.utils.book_new();
    const contexts = [
      { packaging: "彩盒", tax: "含税", sheet: "彩含价-最新" },
      { packaging: "彩盒", tax: "未税", sheet: "彩未价-最新" },
      { packaging: "中性", tax: "含税", sheet: "中含价-最新" },
      { packaging: "中性", tax: "未税", sheet: "中未价-最新" },
    ];
    for (const context of contexts) {
      const source = db.prepare("SELECT normalized_internal, tier, price, attribute FROM prices WHERE packaging=? AND tax=? ORDER BY normalized_internal, id").all(context.packaging, context.tax);
      const grouped = new Map();
      for (const row of source) {
        const item = grouped.get(row.normalized_internal) || { 对外型号: row.normalized_internal, 属性: row.attribute || "" };
        item[row.tier] = mergeTextValues(item[row.tier], row.price);
        item.属性 = mergeTextValues(item.属性, row.attribute);
        grouped.set(row.normalized_internal, item);
      }
      const rows = [...grouped.values()].map((item) => ({ 对外型号: item.对外型号, NP: item.NP ?? "", LP: item.LP ?? "", SLP: item.SLP ?? "", "2S": item["2S"] ?? "", "3S": item["3S"] ?? "", "4S": item["4S"] ?? "", 属性: item.属性 }));
      const worksheet = XLSX.utils.json_to_sheet(rows);
      worksheet["!cols"] = [{ wch: 18 }, ...Array.from({ length: 6 }, () => ({ wch: 12 })), { wch: 14 }];
      XLSX.utils.book_append_sheet(workbook, worksheet, context.sheet);
    }
    return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  }
}

export function exportConfirmedReviews() {
  const rows = db.prepare("SELECT r.*, m.internal FROM reviews r JOIN models m ON m.id=r.model_id WHERE r.status='已确认' ORDER BY r.reviewed_at DESC, r.id DESC").all().map((row) => ({
    对外型号: row.internal,
    OE号: json(row.proposed_oe).join(" / "),
    UF号: json(row.proposed_uf).join(" / "),
    确认时间: row.reviewed_at ? new Date(row.reviewed_at).toLocaleString("zh-CN", { hour12: false }) : "",
  }));
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(rows.length ? rows : [{ 对外型号: "", OE号: "", UF号: "", 确认时间: "" }]);
  worksheet["!cols"] = [{ wch: 20 }, { wch: 40 }, { wch: 28 }, { wch: 22 }];
  XLSX.utils.book_append_sheet(workbook, worksheet, "确认记录");
  return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
}

export function exportPendingReviews() {
  const rows = db.prepare("SELECT r.*, m.internal FROM reviews r LEFT JOIN models m ON m.id=r.model_id WHERE r.status='待校对' ORDER BY r.id DESC").all().map((row) => ({
    查询编号: row.query,
    内部型号: row.internal || "待识别",
    OE号: json(row.proposed_oe).join(" / "),
    UF号: json(row.proposed_uf).join(" / "),
    已验证OE号: json(row.verified_oe).join(" / "),
    已验证UF号: json(row.verified_uf).join(" / "),
    平台: json(row.platforms).join(" / "),
    置信度: row.confidence ?? "",
    查询时间: row.searched_at ? new Date(row.searched_at).toLocaleString("zh-CN", { hour12: false }) : "",
    备注: row.note || "",
  }));
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(rows.length ? rows : [{ 查询编号: "", 内部型号: "", OE号: "", UF号: "", 已验证OE号: "", 已验证UF号: "", 平台: "", 置信度: "", 查询时间: "", 备注: "" }]);
  worksheet["!cols"] = [{ wch: 24 }, { wch: 18 }, { wch: 38 }, { wch: 28 }, { wch: 38 }, { wch: 28 }, { wch: 22 }, { wch: 10 }, { wch: 22 }, { wch: 42 }];
  XLSX.utils.book_append_sheet(workbook, worksheet, "待校对队列");
  return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
}
