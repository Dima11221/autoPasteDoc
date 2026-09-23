"use strict";
const electron = require("electron");
const path = require("node:path");
const fs = require("node:fs");
const Docxtemplater = require("docxtemplater");
const sharp = require("sharp");
const PizZip = require("pizzip");
const IMAGE_EXTENSIONS = /* @__PURE__ */ new Set([".jpg", ".jpeg", ".png"]);
const extractOrderNumber = (fileName) => {
  const baseName = path.parse(fileName).name;
  const match = baseName.match(/\s(\d+)$/);
  if (!match) return null;
  return Number(match[1]);
};
const formatCaption = (photo, captionsByOrder) => {
  const descriptions = captionsByOrder?.get(photo.order);
  if (!descriptions || descriptions.length === 0) {
    return `Фото ${photo.order}.`;
  }
  return `Фото ${photo.order}. ${descriptions.join(" ")}`;
};
const listPhotosSorted = (folderPath) => {
  const entries = fs.readdirSync(folderPath);
  const photos = [];
  const skipped = [];
  for (const fileName of entries) {
    const ext = path.extname(fileName).toLowerCase();
    if (!IMAGE_EXTENSIONS.has(ext)) continue;
    const order = extractOrderNumber(fileName);
    if (order === null) {
      skipped.push(fileName);
      continue;
    }
    photos.push({
      filePath: path.join(folderPath, fileName),
      fileName,
      order
    });
  }
  photos.sort((a, b) => a.order - b.order);
  return { photos, skipped };
};
const LANDSCAPE_WIDTH_CM = 14;
const PORTRAIT_MAX_CM = 10.5;
async function getPhotoSize(filePath) {
  const { data: orientedBuffer, info } = await sharp(filePath).rotate().jpeg({ quality: 90 }).toBuffer({ resolveWithObject: true });
  const widthPx = info.width;
  const heightPx = info.height;
  if (!widthPx || !heightPx) {
    throw new Error(`Не удалось прочитать размеры: ${filePath}`);
  }
  if (widthPx > heightPx) {
    const widthCm = LANDSCAPE_WIDTH_CM;
    const heightCm = widthCm * (heightPx / widthPx);
    return {
      widthCm,
      heightCm,
      widthPx,
      heightPx,
      orientation: "landscape",
      orientedBuffer
    };
  }
  if (heightPx > widthPx) {
    const heightCm = PORTRAIT_MAX_CM;
    const widthCm = heightCm * (widthPx / heightPx);
    return {
      widthCm,
      heightCm,
      widthPx,
      heightPx,
      orientation: "portrait",
      orientedBuffer
    };
  }
  return {
    widthCm: LANDSCAPE_WIDTH_CM,
    heightCm: LANDSCAPE_WIDTH_CM,
    widthPx,
    heightPx,
    orientation: "square",
    orientedBuffer
  };
}
const NO_DEFECTS = "Дефекты и повреждения конструкций не обнаружены";
const HEADING = "Ведомость дефектов и повреждений строительных конструкций";
const parsePhotoNumbers = (cell) => {
  const normalized = cell.replace(/\u00a0/g, " ").trim();
  if (!normalized || normalized === "-" || normalized === "–" || normalized === "—") {
    return [];
  }
  const tokens = normalized.split(/[,;]/).map((t) => t.trim()).filter(Boolean);
  const result = [];
  for (const token of tokens) {
    const range = token.match(/^(\d+)\s*[-–—]\s*(\d+)$/);
    if (range) {
      let from = Number(range[1]);
      let to = Number(range[2]);
      if (from > to) [from, to] = [to, from];
      for (let n = from; n <= to; n++) result.push(n);
      continue;
    }
    if (/^\d+$/.test(token)) {
      result.push(Number(token));
    }
  }
  return result;
};
const cellTextFromXml = (cellXml) => {
  const parts = [];
  const re = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let match;
  while ((match = re.exec(cellXml)) !== null) {
    parts.push(match[1]);
  }
  return parts.join("").replace(/\u00a0/g, " ").trim();
};
const extractTables = (documentXml) => {
  const tables = [];
  const tblRe = /<w:tbl[\s>][\s\S]*?<\/w:tbl>/g;
  let tblMatch;
  while ((tblMatch = tblRe.exec(documentXml)) !== null) {
    const tblXml = tblMatch[0];
    const rows = [];
    const trRe = /<w:tr[\s>][\s\S]*?<\/w:tr>/g;
    let trMatch;
    while ((trMatch = trRe.exec(tblXml)) !== null) {
      const trXml = trMatch[0];
      const cells = [];
      const tcRe = /<w:tc[\s>][\s\S]*?<\/w:tc>/g;
      let tcMatch;
      while ((tcMatch = tcRe.exec(trXml)) !== null) {
        cells.push(cellTextFromXml(tcMatch[0]));
      }
      if (cells.length > 0) rows.push(cells);
    }
    if (rows.length > 0) tables.push(rows);
  }
  return tables;
};
const findDescAndPhotoCols = (header) => {
  let descCol = -1;
  let photoCol = -1;
  for (let i = 0; i < header.length; i++) {
    const h = header[i].toLowerCase();
    if (descCol < 0 && h.includes("описание дефектов")) descCol = i;
    if (photoCol < 0 && h.includes("№ фото")) photoCol = i;
  }
  if (descCol >= 0 && photoCol >= 0) {
    return { descCol, photoCol };
  }
  if (header.length >= 4) {
    return { descCol: 2, photoCol: 3 };
  }
  return null;
};
const pickDefectsTable = (tables) => {
  const matches = [];
  for (const table of tables) {
    const header = table[0] ?? [];
    const joined = header.join(" | ");
    if (joined.includes("№ фото") && joined.toLowerCase().includes("описание дефектов")) {
      matches.push(table);
    }
  }
  if (matches.length > 0) {
    return matches[matches.length - 1];
  }
  return null;
};
const loadCaptionsFromDocx = (templatePath) => {
  const content = fs.readFileSync(templatePath);
  const zip = new PizZip(content);
  const file = zip.file("word/document.xml");
  if (!file) {
    throw new Error("В docx нет word/document.xml");
  }
  const documentXml = file.asText();
  const tables = extractTables(documentXml);
  let table = pickDefectsTable(tables);
  if (!table) {
    const headingPos = documentXml.lastIndexOf(HEADING);
    if (headingPos >= 0) {
      const after = documentXml.slice(headingPos);
      const afterTables = extractTables(after);
      table = afterTables[0] ?? null;
    }
  }
  const captions = /* @__PURE__ */ new Map();
  if (!table || table.length < 2) return captions;
  const cols = findDescAndPhotoCols(table[0]);
  if (!cols) return captions;
  const { descCol, photoCol } = cols;
  for (let r = 1; r < table.length; r++) {
    const row = table[r];
    const description = (row[descCol] ?? "").replace(/\s+/g, " ").trim().replace(/;/g, ".");
    const photoCell = row[photoCol] ?? "";
    if (!description) continue;
    if (description === NO_DEFECTS) continue;
    const numbers = parsePhotoNumbers(photoCell);
    if (numbers.length === 0) continue;
    for (const n of numbers) {
      const list = captions.get(n) ?? [];
      if (!list.includes(description)) list.push(description);
      captions.set(n, list);
    }
  }
  return captions;
};
const paraText = (paraXml) => {
  const parts = [];
  const re = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let match;
  while ((match = re.exec(paraXml)) !== null) {
    parts.push(match[1]);
  }
  return parts.join("").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
};
const listParagraphs = (documentXml) => {
  const result = [];
  const re = /<w:p[\s>][\s\S]*?<\/w:p>/g;
  let match;
  while ((match = re.exec(documentXml)) !== null) {
    result.push({
      start: match.index,
      end: match.index + match[0].length,
      text: paraText(match[0]),
      xml: match[0]
    });
  }
  return result;
};
const isTocPara = (paraXml) => {
  return paraXml.includes("<w:hyperlink") || /w:pStyle\s+[^>]*w:val="21"/.test(paraXml);
};
const isAppendix2Photos = (text) => {
  return /^Приложение №\s*2\b/i.test(text) && /Фотоматериалы/i.test(text);
};
const isAppendix3 = (text) => {
  return /^Приложение №\s*3\b/i.test(text);
};
const DEFAULT_PPR = `<w:pPr>
  <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
  <w:ind w:right="140"/>
  <w:jc w:val="both"/>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
    <w:color w:val="auto"/>
    <w:sz w:val="24"/>
    <w:szCs w:val="24"/>
  </w:rPr>
</w:pPr>`;
const DEFAULT_RPR = `<w:rPr>
  <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
  <w:color w:val="auto"/>
  <w:sz w:val="24"/>
  <w:szCs w:val="24"/>
</w:rPr>`;
const PAGE_BREAK_PARAGRAPH = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
const extractPpr = (paraXml) => {
  const m = paraXml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  return m ? m[0] : null;
};
const tagParagraph = (tag, pPr) => {
  return `<w:p>${pPr}<w:r>${DEFAULT_RPR}<w:t xml:space="preserve">${tag}</w:t></w:r></w:p>`;
};
const withJc = (pPr, jc) => {
  if (/<w:jc\b/.test(pPr)) {
    return pPr.replace(/<w:jc\b[^>]*\/>/, `<w:jc w:val="${jc}"/>`);
  }
  return pPr.replace("</w:pPr>", `<w:jc w:val="${jc}"/></w:pPr>`);
};
const buildPhotosBlock = (imagePpr, captionPpr) => {
  const bothImage = withJc(imagePpr, "both");
  const bothCaption = withJc(captionPpr, "both");
  const leftImage = withJc(imagePpr, "left");
  return tagParagraph("{#photos}", bothImage) + // портрет 10.5 → слева
  tagParagraph("{#isPortrait}", leftImage) + tagParagraph("{%data}", leftImage) + tagParagraph("{caption}", bothCaption) + tagParagraph("{/isPortrait}", leftImage) + // альбом / квадрат → центр
  tagParagraph("{^isPortrait}", bothImage) + tagParagraph("{%data}", bothImage) + tagParagraph("{caption}", bothCaption) + tagParagraph("{/isPortrait}", bothImage) + tagParagraph("{/photos}", bothImage) + PAGE_BREAK_PARAGRAPH;
};
const prepareTemplateZip = (docxBuffer) => {
  const zip = new PizZip(docxBuffer);
  const file = zip.file("word/document.xml");
  if (!file) {
    throw new Error("В docx нет word/document.xml");
  }
  const documentXml = file.asText();
  const paragraphs = listParagraphs(documentXml);
  let startIdx = -1;
  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    if (!isAppendix2Photos(p.text)) continue;
    if (isTocPara(p.xml)) continue;
    startIdx = i;
  }
  if (startIdx < 0) {
    throw new Error("Не найден раздел «Приложение № 2 к Акту обследования. Фотоматериалы»");
  }
  let endIdx = -1;
  for (let i = startIdx + 1; i < paragraphs.length; i++) {
    if (isAppendix3(paragraphs[i].text)) {
      endIdx = i;
      break;
    }
  }
  if (endIdx < 0) {
    throw new Error("Не найден раздел «Приложение № 3 ...» после блока фотоматериалов");
  }
  const between = paragraphs.slice(startIdx + 1, endIdx);
  const imageSample = between.find((p) => p.xml.includes("<w:drawing")) ?? between.find((p) => p.text.includes("{%data}")) ?? between.find((p) => p.text.length > 0);
  const captionSample = between.find((p) => /^Фото\s+\d/i.test(p.text)) ?? between.find((p) => p.text.includes("{caption}")) ?? imageSample;
  const imagePpr = imageSample && extractPpr(imageSample.xml) || DEFAULT_PPR;
  const captionPpr = captionSample && extractPpr(captionSample.xml) || DEFAULT_PPR;
  const photosBlock = buildPhotosBlock(imagePpr, captionPpr);
  const cutFrom = paragraphs[startIdx].end;
  const cutTo = paragraphs[endIdx].start;
  const newXml = documentXml.slice(0, cutFrom) + photosBlock + documentXml.slice(cutTo);
  zip.file("word/document.xml", newXml);
  return zip;
};
const ImageModule = require("docxtemplater-image-module-free");
function cmToModulePx(cm) {
  const EMU_PER_CM = 36e4;
  const EMU_PER_PX = 9525;
  return Math.round(cm * EMU_PER_CM / EMU_PER_PX);
}
async function insertPhotosIntoDocx(options) {
  const { templatePath, photosFolder, outputPath, limit } = options;
  const { photos, skipped } = listPhotosSorted(photosFolder);
  const selected = typeof limit === "number" ? photos.slice(0, limit) : photos;
  if (!fs.existsSync(templatePath)) {
    throw new Error("ENOENT: выбранный Word-файл не найден");
  }
  if (!fs.existsSync(photosFolder)) {
    throw new Error("ENOENT: выбранная папка с фотографиями не найдена");
  }
  if (selected.length === 0) {
    const hint = skipped.length > 0 ? `В папке есть изображения без номера в имени (пропущено: ${skipped.length}). Имя должно заканчиваться на пробел и цифру, например "Фасад 1.jpg"` : "В папке нет файлов .jpg / .jpeg / .png";
    throw new Error(`Нет фото для вставки. ${hint}`);
  }
  const captionsByOrder = loadCaptionsFromDocx(templatePath);
  const sizeCache = /* @__PURE__ */ new Map();
  for (const photo of selected) {
    const size = await getPhotoSize(photo.filePath);
    sizeCache.set(photo.filePath, {
      widthCm: size.widthCm,
      heightCm: size.heightCm,
      orientation: size.orientation,
      orientedBuffer: size.orientedBuffer
    });
  }
  const imageModule = new ImageModule({
    centered: false,
    fileType: "docx",
    getImage: (tagValue) => {
      const cached = sizeCache.get(tagValue);
      if (!cached) {
        throw new Error(`Нет данных картинки: ${tagValue}`);
      }
      return cached.orientedBuffer;
    },
    getSize: (_img, tagValue) => {
      const cached = sizeCache.get(tagValue);
      const widthCm = cached?.widthCm ?? 14;
      const heightCm = cached?.heightCm ?? 14;
      return [cmToModulePx(widthCm), cmToModulePx(heightCm)];
    }
  });
  const content = fs.readFileSync(templatePath);
  const zip = prepareTemplateZip(content);
  const doc = new Docxtemplater(zip, {
    modules: [imageModule],
    paragraphLoop: true,
    linebreaks: true
  });
  doc.render({
    photos: selected.map((photo) => {
      const cached = sizeCache.get(photo.filePath);
      return {
        data: photo.filePath,
        caption: formatCaption(photo, captionsByOrder),
        // портрет = высота 10.5 → выравнивание влево
        isPortrait: cached?.orientation === "portrait"
      };
    })
  });
  const buffer = doc.toBuffer();
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, buffer);
  return {
    outputPath,
    insertedCount: selected.length,
    skipped
  };
}
const humanizeError = (error) => {
  const raw = error instanceof Error ? error.message : String(error);
  const lower = raw.toLowerCase();
  if (lower.includes("ebusy") || lower.includes("eperm") || lower.includes("eacces") || lower.includes("resource busy") || lower.includes("operation not permitted")) {
    return "Не удалось сохранить файл. Закройте документ в Word (если открыт) и попробуйте снова.";
  }
  if (lower.includes("enoent")) {
    return "Не удалось найти файл шаблона. Проверьте путь к файлу и попробуйте снова.";
  }
  if (lower.includes("приложение № 2") || lower.includes("приложение № 3") || lower.includes("фотоматериалы")) {
    return "Не найден раздел для вставки фото (Приложение № 2 … Фотоматериалы → Приложение № 3).";
  }
  if (lower.includes("multierror") || lower.includes("unopened_tag") || lower.includes("unclosed_tag") || lower.includes("duplicate_open_tag") || lower.includes("closing_tag_does_not_match") || lower.includes("xmltemplater") || lower.includes('tag "photos"') || lower.includes("photos") && lower.includes("not found")) {
    return "В документе не найден корректный блок для фото. Нужны метки \n{#photos}\n{#data}\n\n{/photos}.";
  }
  if (lower.includes("zip") || lower.includes("end of central directory")) {
    return "Не удалось прочитать документ. Убедитесь, что выбран именно .docx (не .doc) и файл не поврежден.";
  }
  return raw;
};
function createWindow() {
  const win = new electron.BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, "../preload/index.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });
  if (process.env.ELECTRON_RENDERER_URL) {
    win.loadURL(process.env.ELECTRON_RENDERER_URL);
  } else {
    win.loadFile(path.join(__dirname, "../renderer/index.html"));
  }
}
function buildOutputPath(templatePath) {
  const { dir, name, ext } = path.parse(templatePath);
  return path.join(dir, `${name}_с_фото${ext}`);
}
function registerIpc() {
  electron.ipcMain.handle("dialog:selectDocx", async () => {
    const result = await electron.dialog.showOpenDialog({
      title: "Выберите Word-документ",
      properties: ["openFile"],
      filters: [{ name: "Word", extensions: ["docx"] }]
    });
    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths[0];
  });
  electron.ipcMain.handle("dialog:selectPhotosFolder", async () => {
    const result = await electron.dialog.showOpenDialog({
      title: "Выберите папку с фото",
      properties: ["openDirectory"]
    });
    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths[0];
  });
  electron.ipcMain.handle(
    "insert:run",
    async (_event, payload) => {
      try {
        const outputPath = buildOutputPath(payload.templatePath);
        const result = await insertPhotosIntoDocx({
          templatePath: payload.templatePath,
          photosFolder: payload.photosFolder,
          outputPath
        });
        return {
          ok: true,
          outputPath: result.outputPath,
          insertedCount: result.insertedCount,
          skipped: result.skipped
        };
      } catch (error) {
        return { ok: false, error: humanizeError(error) };
      }
    }
  );
  electron.ipcMain.handle("shell:showItemInFolder", async (_event, filePath) => {
    if (!filePath) return false;
    electron.shell.showItemInFolder(filePath);
    return true;
  });
}
electron.app.whenReady().then(() => {
  registerIpc();
  createWindow();
  electron.app.on("activate", () => {
    if (electron.BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});
electron.app.on("window-all-closed", () => {
  if (process.platform !== "darwin") electron.app.quit();
});
