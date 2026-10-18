import * as XLSX from "xlsx";
import { extractTextFromFile, extensionOf, isSupportedQuoteFile } from "./quote-file-parser";
import { prepareProductCatalogFile } from "./xlsx-images";

function cleanToken(value = "") {
  return String(value).toUpperCase().replace(/^[^A-Z0-9]+|[^A-Z0-9]+$/g, "").replace(/[._\\]+/g, "-");
}

function identifiers(line = "") {
  return [...line.matchAll(/[A-Z0-9][A-Z0-9._/\\-]{3,25}/gi)]
    .map((match) => cleanToken(match[0]))
    .filter((value) => /\d/.test(value) && !/^(?:19|20)\d{2}$/.test(value) && !/^\d{1,4}$/.test(value));
}

function createWorkbookFile(rows, sheetName, file) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), sheetName.slice(0, 31));
  const output = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
  return new File([output], `${file.name.replace(/\.[^.]+$/, "")}.xlsx`, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function parseModelRows(text) {
  const rows = [];
  for (const rawLine of String(text).split(/\r?\n/)) {
    const values = [...new Set(identifiers(rawLine))];
    if (!values.length) continue;
    const internal = values.find((value) => /^DQ[A-Z0-9-]*\d/i.test(value)) || values.find((value) => !/^UF/i.test(value));
    if (!internal) continue;
    const uf = values.filter((value) => /^UF/i.test(value));
    const oe = values.filter((value) => value !== internal && !/^UF/i.test(value));
    rows.push({ 对外型号: internal, OE号: oe.join(" / "), UF号: uf.join(" / "), 识别原文: rawLine.trim() });
  }
  return rows;
}

function priceContext(text) {
  if (/彩(?:盒)?[^\n]{0,8}未税|彩未价|彩盒未税/i.test(text)) return { sheet: "彩未价-识别" };
  if (/彩(?:盒)?[^\n]{0,8}含税|彩含价|彩盒含税/i.test(text)) return { sheet: "彩含价-识别" };
  if (/中(?:性)?[^\n]{0,8}未税|中未价|中性未税/i.test(text)) return { sheet: "中未价-识别" };
  if (/中(?:性)?[^\n]{0,8}含税|中含价|中性含税/i.test(text)) return { sheet: "中含价-识别" };
  return null;
}

function parsePriceRows(text) {
  const rows = [];
  for (const rawLine of String(text).split(/\r?\n/)) {
    const model = identifiers(rawLine).find((value) => /^DQ/i.test(value));
    if (!model) continue;
    const tail = rawLine.slice(rawLine.toUpperCase().indexOf(model) + model.length);
    const prices = [...tail.matchAll(/(?:^|\s)(\d+(?:\.\d+)?)(?=\s|$)/g)].map((match) => Number(match[1])).filter(Number.isFinite);
    if (prices.length < 4) continue;
    const tiers = ["NP", "LP", "SLP", "2S", "3S", "4S"];
    rows.push(Object.fromEntries([["对外型号", model], ...tiers.map((tier, index) => [tier, prices[index] ?? ""]) ]));
  }
  return rows;
}

function catalogJsonFile(records, file) {
  return new File([JSON.stringify({ filename: file.name, records })], `${file.name}.catalog.json`, { type: "application/json" });
}

function parseCatalogTextRows(text, sourceName = "文字识别") {
  const records = [];
  String(text).split(/\r?\n/).forEach((rawLine, index) => {
    const values = [...new Set(identifiers(rawLine))];
    const internal = values.find((value) => /^DQ[A-Z0-9-]*\d/i.test(value));
    if (!internal) return;
    const oe = values.filter((value) => value !== internal && !/^UF/i.test(value));
    records.push({ internal, normalizedInternal: cleanToken(internal).replace(/-/g, ""), commonName: "", vehicle: rawLine.trim(), oe: oe.join(" / "), note: "文字识别导入", imageData: "", sourceSheet: sourceName, sourceRow: index + 1 });
  });
  return records;
}

async function prepareLegacyCatalogFile(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", raw: false });
  const records = [];
  workbook.SheetNames.forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: false });
    const headerIndex = rows.findIndex((row) => row.some((cell) => /画册型号|对外型号|内部型号/i.test(String(cell))));
    if (headerIndex < 0) return;
    const headers = rows[headerIndex].map((cell) => String(cell).trim());
    const indexOf = (pattern) => headers.findIndex((header) => pattern.test(header));
    const internalCol = indexOf(/画册型号|对外型号|内部型号/i);
    const nameCol = indexOf(/通用名|产品名称/i);
    const vehicleCol = indexOf(/适用车型|车型/i);
    const oeCol = indexOf(/OE/i);
    const noteCol = indexOf(/备注/i);
    rows.slice(headerIndex + 1).forEach((row, offset) => {
      const internal = String(row[internalCol] || "").replace(/\s+/g, "").trim();
      if (!internal || !/\d/.test(internal)) return;
      records.push({ internal, normalizedInternal: internal.toUpperCase().replace(/[^A-Z0-9]/g, ""), commonName: String(row[nameCol] || "").trim(), vehicle: String(row[vehicleCol] || "").trim(), oe: String(row[oeCol] || "").trim(), note: String(row[noteCol] || "").trim(), imageData: "", sourceSheet: sheetName, sourceRow: headerIndex + offset + 2 });
    });
  });
  if (!records.length) throw new Error("产品目录中没有识别到型号记录");
  return catalogJsonFile(records, file);
}

export async function prepareKnowledgeFile(file, type, onProgress) {
  if (!isSupportedQuoteFile(file)) throw new Error("仅支持文字、Excel、PDF、PNG 和 JPG 文件");
  if (type === "catalog") {
    if (extensionOf(file.name) === "xlsx") return prepareProductCatalogFile(file, onProgress);
    if (extensionOf(file.name) === "xls") return prepareLegacyCatalogFile(file);
    const text = await extractTextFromFile(file, onProgress);
    const records = parseCatalogTextRows(text, file.name);
    if (!records.length) throw new Error("没有识别到以 DQ 开头的产品目录型号");
    return catalogJsonFile(records, file);
  }
  if (type === "relations") {
    if (!["xlsx", "xls"].includes(extensionOf(file.name))) throw new Error("关系核验表仅支持 Excel 文件");
    return file;
  }
  if (["xlsx", "xls"].includes(extensionOf(file.name))) return file;
  const text = await extractTextFromFile(file, onProgress);
  if (type === "models") {
    const rows = parseModelRows(text);
    if (!rows.length) throw new Error("没有识别到可用的内部型号 / OE / UF 关系");
    return createWorkbookFile(rows, "型号匹配", file);
  }
  const context = priceContext(text);
  if (!context) throw new Error("未识别出彩盒/中性及含税/未税，请在文件文字中标明价格类型");
  const rows = parsePriceRows(text);
  if (!rows.length) throw new Error("没有识别到型号及对应价格列");
  return createWorkbookFile(rows, context.sheet, file);
}
