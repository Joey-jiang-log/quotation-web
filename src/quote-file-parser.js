import * as XLSX from "xlsx";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { extractInquiryRows } from "./inquiry-text-parser";
import { extractExcelRowImages } from "./xlsx-images";

export { extractInquiryRows } from "./inquiry-text-parser";

const SUPPORTED_EXTENSIONS = ["xlsx", "xls", "csv", "txt", "pdf", "png", "jpg", "jpeg"];

export function extensionOf(name = "") {
  return String(name).split(".").pop()?.toLowerCase() || "";
}

export function isSupportedQuoteFile(file) {
  return Boolean(file && SUPPORTED_EXTENSIONS.includes(extensionOf(file.name)));
}

export function quoteFileType(file) {
  const extension = extensionOf(file?.name);
  if (["xlsx", "xls"].includes(extension)) return "Excel";
  if (["csv", "txt"].includes(extension)) return "文字";
  if (extension === "pdf") return "PDF";
  if (["png", "jpg", "jpeg"].includes(extension)) return "图片";
  return "未知格式";
}

function rowsToWorkbookFile(rows, originalFile) {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(rows);
  worksheet["!cols"] = [{ wch: 24 }, { wch: 10 }, { wch: 68 }];
  XLSX.utils.book_append_sheet(workbook, worksheet, "识别结果");
  const output = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
  return new File([output], originalFile.name, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function groupPdfTextItems(items = []) {
  const lines = [];
  const sorted = items
    .filter((item) => String(item.str || "").trim())
    .map((item) => ({ text: String(item.str).trim(), x: Number(item.transform?.[4] || 0), y: Number(item.transform?.[5] || 0) }))
    .sort((a, b) => Math.abs(b.y - a.y) > 3 ? b.y - a.y : a.x - b.x);
  for (const item of sorted) {
    const line = lines.find((entry) => Math.abs(entry.y - item.y) <= 3);
    if (line) line.items.push(item);
    else lines.push({ y: item.y, items: [item] });
  }
  return lines
    .sort((a, b) => b.y - a.y)
    .map((line) => line.items.sort((a, b) => a.x - b.x).map((item) => item.text).join(" "))
    .join("\n");
}

async function createOcrWorker(onProgress) {
  const { createWorker } = await import("tesseract.js");
  const worker = await createWorker("eng", 1, {
    logger: (message) => {
      if (message.status === "recognizing text") onProgress?.({ stage: "正在识别图片文字", progress: Math.round((message.progress || 0) * 100) });
    },
  });
  await worker.setParameters({ tessedit_pageseg_mode: "6", preserve_interword_spaces: "1" });
  return worker;
}

async function preprocessImage(source) {
  const sourceCanvas = window.document.createElement("canvas");
  let width;
  let height;
  if (source instanceof HTMLCanvasElement) {
    width = source.width;
    height = source.height;
    sourceCanvas.width = width;
    sourceCanvas.height = height;
    sourceCanvas.getContext("2d", { willReadFrequently: true }).drawImage(source, 0, 0);
  } else {
    const bitmap = await createImageBitmap(source);
    width = bitmap.width;
    height = bitmap.height;
    sourceCanvas.width = width;
    sourceCanvas.height = height;
    sourceCanvas.getContext("2d", { willReadFrequently: true }).drawImage(bitmap, 0, 0);
    bitmap.close();
  }

  const context = sourceCanvas.getContext("2d", { willReadFrequently: true });
  const image = context.getImageData(0, 0, width, height);
  const pixels = image.data;
  const darkRows = new Uint32Array(height);
  const darkColumns = new Uint32Array(width);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const index = (y * width + x) * 4;
      const gray = pixels[index] * .299 + pixels[index + 1] * .587 + pixels[index + 2] * .114;
      const binary = gray < 172 ? 0 : 255;
      pixels[index] = binary;
      pixels[index + 1] = binary;
      pixels[index + 2] = binary;
      if (binary === 0) { darkRows[y] += 1; darkColumns[x] += 1; }
    }
  }
  const removeRows = [...darkRows].map((count, index) => count > width * .45 ? index : -1).filter((index) => index >= 0);
  const removeColumns = [...darkColumns].map((count, index) => count > height * .45 ? index : -1).filter((index) => index >= 0);
  const whiten = (x, y) => {
    const index = (y * width + x) * 4;
    pixels[index] = 255; pixels[index + 1] = 255; pixels[index + 2] = 255; pixels[index + 3] = 255;
  };
  for (const y of removeRows) for (let offset = -1; offset <= 1; offset += 1) if (y + offset >= 0 && y + offset < height) for (let x = 0; x < width; x += 1) whiten(x, y + offset);
  for (const x of removeColumns) for (let offset = -1; offset <= 1; offset += 1) if (x + offset >= 0 && x + offset < width) for (let y = 0; y < height; y += 1) whiten(x + offset, y);
  context.putImageData(image, 0, 0);

  const scale = Math.min(3, Math.max(1, 3000 / Math.max(1, width)));
  if (scale === 1) return sourceCanvas;
  const output = window.document.createElement("canvas");
  output.width = Math.round(width * scale);
  output.height = Math.round(height * scale);
  const outputContext = output.getContext("2d", { willReadFrequently: true });
  outputContext.imageSmoothingEnabled = true;
  outputContext.imageSmoothingQuality = "high";
  outputContext.drawImage(sourceCanvas, 0, 0, output.width, output.height);
  return output;
}

async function ocrImage(file, onProgress) {
  onProgress?.({ stage: "正在优化图片清晰度", progress: 5 });
  const preparedImage = await preprocessImage(file);
  const worker = await createOcrWorker(onProgress);
  try {
    const result = await worker.recognize(preparedImage);
    return result.data.text || "";
  } finally {
    await worker.terminate();
  }
}

function groupedLineIndexes(counts, threshold) {
  const groups = [];
  counts.forEach((count, index) => {
    if (count < threshold) return;
    const previous = groups[groups.length - 1];
    if (previous && index <= previous[1] + 2) previous[1] = index;
    else groups.push([index, index]);
  });
  return groups.map(([start, end]) => Math.round((start + end) / 2));
}

function cropCanvasDataUrl(source, top, bottom) {
  const cropHeight = Math.max(1, bottom - top);
  const width = 180;
  const height = 120;
  const scale = Math.min(width / source.width, height / cropHeight);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  context.fillStyle = "#fff";
  context.fillRect(0, 0, width, height);
  const targetWidth = Math.max(1, Math.round(source.width * scale));
  const targetHeight = Math.max(1, Math.round(cropHeight * scale));
  context.drawImage(source, 0, top, source.width, cropHeight, Math.round((width - targetWidth) / 2), Math.round((height - targetHeight) / 2), targetWidth, targetHeight);
  return canvas.toDataURL("image/jpeg", .72);
}

function detectTableRowCrops(canvas, limit) {
  const context = canvas.getContext("2d", { willReadFrequently: true });
  const { data } = context.getImageData(0, 0, canvas.width, canvas.height);
  const darkRows = new Uint32Array(canvas.height);
  for (let y = 0; y < canvas.height; y += 1) {
    for (let x = 0; x < canvas.width; x += 2) {
      const index = (y * canvas.width + x) * 4;
      const gray = data[index] * .299 + data[index + 1] * .587 + data[index + 2] * .114;
      if (gray < 115) darkRows[y] += 1;
    }
  }
  const lines = groupedLineIndexes([...darkRows], canvas.width * .16);
  const regions = [];
  for (let index = 0; index < lines.length - 1; index += 1) {
    if (lines[index + 1] - lines[index] >= 30) regions.push([lines[index] + 2, lines[index + 1] - 2]);
  }
  return regions.slice(-limit).map(([top, bottom]) => cropCanvasDataUrl(canvas, top, bottom));
}

async function sourceToCanvas(source, maxWidth = 1500) {
  const bitmap = await createImageBitmap(source);
  const scale = Math.min(1, maxWidth / bitmap.width);
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(bitmap.width * scale));
  canvas.height = Math.max(1, Math.round(bitmap.height * scale));
  canvas.getContext("2d", { willReadFrequently: true }).drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  bitmap.close();
  return canvas;
}

async function extractVisualRowImages(file, rowCount, onProgress) {
  if (!rowCount) return { images: [], status: "" };
  const extension = extensionOf(file.name);
  const images = [];
  if (["png", "jpg", "jpeg"].includes(extension)) {
    const canvas = await sourceToCanvas(file);
    if (rowCount === 1) return { images: [cropCanvasDataUrl(canvas, 0, canvas.height)], status: "" };
    images.push(...detectTableRowCrops(canvas, rowCount));
  } else if (extension === "pdf") {
    const pdfjs = await import("pdfjs-dist");
    pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
    const pdfDocument = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
    for (let pageNumber = 1; pageNumber <= Math.min(pdfDocument.numPages, 8) && images.length < rowCount; pageNumber += 1) {
      const page = await pdfDocument.getPage(pageNumber);
      const base = page.getViewport({ scale: 1 });
      const viewport = page.getViewport({ scale: Math.min(2, 1500 / Math.max(1, base.width)) });
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      await page.render({ canvasContext: canvas.getContext("2d", { willReadFrequently: true }), viewport }).promise;
      images.push(...detectTableRowCrops(canvas, rowCount - images.length));
      onProgress?.({ stage: `正在按表格行切分 PDF 图片 ${pageNumber}/${Math.min(pdfDocument.numPages, 8)}`, progress: 92 });
    }
  }
  return { images: images.slice(0, rowCount), status: images.length >= rowCount ? "" : "图片待确认" };
}

async function parsePdf(file, onProgress) {
  const pdfjs = await import("pdfjs-dist");
  pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
  const pdfDocument = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
  const pageCount = Math.min(pdfDocument.numPages, 8);
  const extractedPages = [];
  onProgress?.({ stage: "正在读取 PDF 文字", progress: 8 });
  for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
    const page = await pdfDocument.getPage(pageNumber);
    const textContent = await page.getTextContent();
    extractedPages.push(groupPdfTextItems(textContent.items));
    onProgress?.({ stage: "正在读取 PDF 文字", progress: Math.round(8 + (pageNumber / pageCount) * 27) });
  }
  let text = extractedPages.join("\n");
  if (extractInquiryRows(text).length) return text;

  onProgress?.({ stage: "PDF 为扫描件，正在启动文字识别", progress: 38 });
  const worker = await createOcrWorker((message) => onProgress?.(message));
  const ocrPages = [];
  try {
    for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
      const page = await pdfDocument.getPage(pageNumber);
      const baseViewport = page.getViewport({ scale: 1 });
      const scale = Math.min(2.2, 2200 / Math.max(1, baseViewport.width));
      const viewport = page.getViewport({ scale });
      const canvas = window.document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      await page.render({ canvasContext: canvas.getContext("2d", { willReadFrequently: true }), viewport }).promise;
      const result = await worker.recognize(await preprocessImage(canvas));
      ocrPages.push(result.data.text || "");
      onProgress?.({ stage: `正在识别 PDF 第 ${pageNumber}/${pageCount} 页`, progress: Math.round(40 + (pageNumber / pageCount) * 55) });
    }
  } finally {
    await worker.terminate();
  }
  text = ocrPages.join("\n");
  return text;
}

export async function extractTextFromFile(file, onProgress) {
  if (!isSupportedQuoteFile(file)) throw new Error("仅支持 .xlsx、.xls、.csv、.txt、.pdf、.png、.jpg 文件");
  if (file.size > 25 * 1024 * 1024) throw new Error("文件不能超过 25MB");
  const extension = extensionOf(file.name);
  if (["csv", "txt"].includes(extension)) {
    onProgress?.({ stage: "正在读取文字内容", progress: 30 });
    return file.text();
  }
  if (extension === "pdf") return parsePdf(file, onProgress);
  if (["png", "jpg", "jpeg"].includes(extension)) return ocrImage(file, onProgress);
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", raw: false });
  return workbook.SheetNames.map((name) => XLSX.utils.sheet_to_csv(workbook.Sheets[name])).join("\n");
}

export async function prepareQuoteFile(file, onProgress) {
  if (!isSupportedQuoteFile(file)) throw new Error("仅支持 .xlsx、.xls、.csv、.txt、.pdf、.png、.jpg 文件");
  if (file.size > 25 * 1024 * 1024) throw new Error("文件不能超过 25MB");
  const extension = extensionOf(file.name);
  if (["xlsx", "xls"].includes(extension)) {
    const customerImages = extension === "xlsx" ? await extractExcelRowImages(file, onProgress) : {};
    return { file, extractedRows: null, mode: "Excel", customerImages, imageValidation: "embedded-image-column", imageStatus: "" };
  }

  onProgress?.({ stage: extension === "pdf" ? "正在读取 PDF" : ["csv", "txt"].includes(extension) ? "正在读取文字" : "正在读取图片", progress: 3 });
  const text = await extractTextFromFile(file, onProgress);
  const rows = extractInquiryRows(text);
  if (!rows.length) throw new Error("没有识别到型号 / OE 内容，请上传更清晰的文件或改用 Excel");
  onProgress?.({ stage: `已识别 ${rows.length} 条型号，正在匹配价格`, progress: 100 });
  return {
    file: rowsToWorkbookFile(rows, file),
    extractedRows: rows.length,
    mode: extension === "pdf" ? "PDF 识别" : ["csv", "txt"].includes(extension) ? "文字解析" : "图片识别",
    customerImagesByIndex: [],
    imageValidation: "",
    imageStatus: "",
  };
}
