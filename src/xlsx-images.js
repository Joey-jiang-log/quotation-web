import JSZip from "jszip";

const CATALOG_SHEETS = ["已量产目录", "客户专用型号目录", "20260625新增13款"];

function parseXml(text) {
  const document = new DOMParser().parseFromString(text, "application/xml");
  if (document.querySelector("parsererror")) throw new Error("Excel 文件结构无法解析");
  return document;
}

function nodes(parent, localName) {
  return Array.from(parent.getElementsByTagNameNS("*", localName));
}

function first(parent, localName) {
  return nodes(parent, localName)[0] || null;
}

function attribute(node, localName) {
  return Array.from(node?.attributes || []).find((item) => item.localName === localName)?.value || "";
}

function dirname(path) {
  return path.split("/").slice(0, -1).join("/");
}

function normalizePath(path) {
  const output = [];
  for (const item of path.split("/")) {
    if (!item || item === ".") continue;
    if (item === "..") output.pop();
    else output.push(item);
  }
  return output.join("/");
}

function resolvePath(source, target) {
  return target.startsWith("/") ? target.slice(1) : normalizePath(`${dirname(source)}/${target}`);
}

function relsPath(source) {
  const parts = source.split("/");
  const file = parts.pop();
  return `${parts.join("/")}/_rels/${file}.rels`;
}

async function relationshipMap(zip, source) {
  const entry = zip.file(relsPath(source));
  if (!entry) return new Map();
  const document = parseXml(await entry.async("text"));
  return new Map(nodes(document, "Relationship").map((item) => [item.getAttribute("Id"), {
    target: resolvePath(source, item.getAttribute("Target") || ""),
    type: item.getAttribute("Type") || "",
  }]));
}

async function workbookSheets(zip) {
  const workbook = parseXml(await zip.file("xl/workbook.xml").async("text"));
  const rels = await relationshipMap(zip, "xl/workbook.xml");
  const result = new Map();
  for (const sheet of nodes(workbook, "sheet")) {
    const relation = rels.get(attribute(sheet, "id"));
    if (relation) result.set(sheet.getAttribute("name"), relation.target);
  }
  return result;
}

async function drawingRows(zip, sheetPath) {
  const sheetRels = await relationshipMap(zip, sheetPath);
  const drawingPath = [...sheetRels.values()].find((item) => item.type.endsWith("/drawing"))?.target;
  if (!drawingPath || !zip.file(drawingPath)) return new Map();
  const drawing = parseXml(await zip.file(drawingPath).async("text"));
  const drawingRels = await relationshipMap(zip, drawingPath);
  const result = new Map();
  for (const anchor of [...nodes(drawing, "twoCellAnchor"), ...nodes(drawing, "oneCellAnchor")]) {
    const origin = first(anchor, "from");
    const row = Number(first(origin, "row")?.textContent);
    const col = Number(first(origin, "col")?.textContent);
    const relation = drawingRels.get(attribute(first(anchor, "blip"), "embed"));
    if (Number.isFinite(row) && relation?.target) {
      const current = result.get(row) || [];
      current.push({ col, path: relation.target });
      result.set(row, current);
    }
  }
  return result;
}

function blobToImage(blob) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => { URL.revokeObjectURL(url); resolve(image); };
    image.onerror = () => { URL.revokeObjectURL(url); reject(new Error("图片读取失败")); };
    image.src = url;
  });
}

function isLikelyProductImage(image) {
  const ratio = image.naturalWidth / Math.max(1, image.naturalHeight);
  if (ratio > 4.2 || ratio < .18) return false;
  const sample = document.createElement("canvas");
  const scale = Math.min(1, 260 / Math.max(image.naturalWidth, image.naturalHeight));
  sample.width = Math.max(1, Math.round(image.naturalWidth * scale));
  sample.height = Math.max(1, Math.round(image.naturalHeight * scale));
  const context = sample.getContext("2d", { willReadFrequently: true });
  context.fillStyle = "#fff";
  context.fillRect(0, 0, sample.width, sample.height);
  context.drawImage(image, 0, 0, sample.width, sample.height);
  const pixels = context.getImageData(0, 0, sample.width, sample.height).data;
  let foreground = 0;
  for (let index = 0; index < pixels.length; index += 4) {
    if (pixels[index] < 242 || pixels[index + 1] < 242 || pixels[index + 2] < 242) foreground += 1;
  }
  return foreground / Math.max(1, pixels.length / 4) >= .025;
}

async function compactImage(blob, width = 180, height = 130, requireProductImage = false) {
  const image = await blobToImage(blob);
  if (requireProductImage && !isLikelyProductImage(image)) return "";
  const scale = Math.min(width / image.naturalWidth, height / image.naturalHeight, 1);
  const drawWidth = Math.max(1, Math.round(image.naturalWidth * scale));
  const drawHeight = Math.max(1, Math.round(image.naturalHeight * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.drawImage(image, Math.round((width - drawWidth) / 2), Math.round((height - drawHeight) / 2), drawWidth, drawHeight);
  return canvas.toDataURL("image/jpeg", .72);
}

async function sharedStrings(zip) {
  const entry = zip.file("xl/sharedStrings.xml");
  if (!entry) return [];
  const document = parseXml(await entry.async("text"));
  return nodes(document, "si").map((item) => nodes(item, "t").map((text) => text.textContent || "").join(""));
}

function columnName(reference = "") {
  return reference.match(/^[A-Z]+/i)?.[0]?.toUpperCase() || "";
}

function sheetRows(document, strings) {
  const result = new Map();
  for (const row of nodes(document, "row")) {
    const rowNumber = Number(row.getAttribute("r"));
    const values = {};
    for (const cell of nodes(row, "c")) {
      const type = cell.getAttribute("t");
      const raw = first(cell, "v")?.textContent || "";
      const value = type === "s" ? strings[Number(raw)] || "" : type === "inlineStr" ? nodes(cell, "t").map((item) => item.textContent || "").join("") : raw;
      values[columnName(cell.getAttribute("r"))] = value;
    }
    result.set(rowNumber, values);
  }
  return result;
}

function normalizeModel(value = "") {
  return String(value).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function columnIndex(name = "") {
  return String(name).toUpperCase().split("").reduce((value, character) => value * 26 + character.charCodeAt(0) - 64, 0) - 1;
}

function inquiryImageColumns(rows) {
  const result = new Set();
  for (const [rowNumber, values] of rows) {
    if (rowNumber > 20) break;
    for (const [column, rawValue] of Object.entries(values)) {
      const value = String(rawValue || "").trim().toLowerCase().replace(/\s+/g, "");
      if (/^(客户)?(询价)?(产品)?(实物)?图片$/.test(value) || /^(image|photo|picture|productimage|productphoto)$/.test(value)) {
        result.add(columnIndex(column));
      }
    }
  }
  return result;
}

export async function extractExcelRowImages(file, onProgress) {
  if (!/\.xlsx$/i.test(file.name)) return {};
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const sheets = await workbookSheets(zip);
  const strings = await sharedStrings(zip);
  const result = {};
  let processed = 0;
  for (const [sheetName, sheetPath] of sheets) {
    const sheetEntry = zip.file(sheetPath);
    if (!sheetEntry) continue;
    const sheetValues = sheetRows(parseXml(await sheetEntry.async("text")), strings);
    const imageColumns = inquiryImageColumns(sheetValues);
    if (!imageColumns.size) continue;
    const rows = await drawingRows(zip, sheetPath);
    const rowImages = {};
    for (const [row, images] of rows) {
      const selected = [...images].filter((item) => imageColumns.has(item.col)).sort((a, b) => a.col - b.col)[0];
      const entry = selected && zip.file(selected.path);
      if (!entry || processed >= 200) continue;
      try {
        const image = await compactImage(await entry.async("blob"), 180, 130, true);
        if (!image) continue;
        rowImages[row] = image;
        processed += 1;
        onProgress?.({ stage: `正在读取 Excel 产品图片 ${processed}`, progress: Math.min(90, 10 + processed) });
      } catch { /* 单张异常不影响报价 */ }
    }
    if (Object.keys(rowImages).length) result[sheetName] = rowImages;
  }
  return result;
}

export async function prepareProductCatalogFile(file, onProgress) {
  if (!/\.xlsx$/i.test(file.name)) throw new Error("产品目录需上传 .xlsx 文件");
  onProgress?.({ stage: "正在读取产品目录结构", progress: 2 });
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const sheets = await workbookSheets(zip);
  const strings = await sharedStrings(zip);
  const records = new Map();
  let imageCount = 0;
  for (let sheetIndex = 0; sheetIndex < CATALOG_SHEETS.length; sheetIndex += 1) {
    const sheetName = CATALOG_SHEETS[sheetIndex];
    const sheetPath = sheets.get(sheetName);
    if (!sheetPath || !zip.file(sheetPath)) continue;
    const rows = sheetRows(parseXml(await zip.file(sheetPath).async("text")), strings);
    const images = await drawingRows(zip, sheetPath);
    for (const [rowNumber, values] of rows) {
      if (rowNumber < 4) continue;
      const internal = String(values.B || "").replace(/\s+/g, "").trim();
      const key = normalizeModel(internal);
      if (!key) continue;
      const selected = (images.get(rowNumber - 1) || []).find((item) => item.col === 2) || (images.get(rowNumber - 1) || [])[0];
      let imageData = "";
      if (selected && zip.file(selected.path)) {
        try {
          imageData = await compactImage(await zip.file(selected.path).async("blob"), 220, 150);
          imageCount += 1;
          onProgress?.({ stage: `正在压缩产品图片 ${imageCount}`, progress: Math.min(96, 8 + Math.round(imageCount / 8)) });
        } catch { /* 保留无图记录 */ }
      }
      records.set(key, {
        internal,
        normalizedInternal: key,
        commonName: String(values.D || "").trim(),
        vehicle: String(values.E || "").trim(),
        oe: String(values.F || "").trim(),
        note: String(sheetName === "已量产目录" ? values.H || "" : values.G || "").trim(),
        imageData,
        sourceSheet: sheetName,
        sourceRow: rowNumber,
        priority: sheetIndex + 1,
      });
    }
  }
  if (!records.size) throw new Error("三个产品目录工作表中没有识别到画册型号");
  onProgress?.({ stage: `目录解析完成，共 ${records.size} 个型号`, progress: 100 });
  return new File([JSON.stringify({ filename: file.name, visibleSheets: CATALOG_SHEETS, records: [...records.values()] })], `${file.name}.catalog.json`, { type: "application/json" });
}
