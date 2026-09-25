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
const FIELD_DEFS = [
  { key: "yearBuilt", group: "general", sourceLabel: "Год постройки (ввода в эксплуатацию)", formLabel: "Год постройки" },
  { key: "buildingArea", group: "general", sourceLabel: "Площадь застройки", formLabel: "Площадь застройки" },
  { key: "buildingVolume", group: "general", sourceLabel: "Строительный объем", formLabel: "Строительный объем" },
  { key: "floors", group: "general", sourceLabel: "Количество этажей", formLabel: "Количество этажей" },
  { key: "fireCategory", group: "general", sourceLabel: "Категория помещений по пожарной", formLabel: "Категория по пожарной опасности" },
  { key: "responsibilityLevel", group: "general", sourceLabel: "Уровень ответственности", formLabel: "Уровень ответственности" },
  { key: "snowRegion", group: "general", sourceLabel: "Снеговой район", formLabel: "Снеговой район" },
  { key: "windRegion", group: "general", sourceLabel: "Ветровой район", formLabel: "Ветровой район" },
  { key: "seismicity", group: "general", sourceLabel: "Сейсмичность района", formLabel: "Сейсмичность" },
  { key: "climateRegion", group: "general", sourceLabel: "Климатический район", formLabel: "Климатический район" },
  { key: "designTemp", group: "general", sourceLabel: "Расчетная температура наружного воздуха", formLabel: "Расчётная температура" },
  { key: "volumePlanning", group: "construction", sourceLabel: "Объемно-планировочные решения", formLabel: "Объёмно-планировочные решения" },
  { key: "structuralScheme", group: "construction", sourceLabel: "Конструктивная схема", formLabel: "Конструктивная схема" },
  { key: "foundations", group: "construction", sourceLabel: "Фундаменты", formLabel: "Фундаменты" },
  { key: "outerWalls", group: "construction", sourceLabel: "Наружные стены", formLabel: "Наружные стены" },
  { key: "innerWalls", group: "construction", sourceLabel: "Внутренние стены и перегородки", formLabel: "Внутренние стены и перегородки" },
  { key: "columns", group: "construction", sourceLabel: "Колонны", formLabel: "Колонны" },
  { key: "covering", group: "construction", sourceLabel: "Покрытие", formLabel: "Покрытие" },
  { key: "floorSlab", group: "construction", sourceLabel: "Перекрытие", formLabel: "Перекрытие" },
  { key: "roof", group: "construction", sourceLabel: "Кровля, водосточная система", formLabel: "Кровля, водосток" },
  { key: "openings", group: "construction", sourceLabel: "Ворота, оконные и дверные заполнения", formLabel: "Ворота / окна / двери" },
  { key: "floorsDrainage", group: "construction", sourceLabel: "Полы и дренажная система", formLabel: "Полы и дренаж" },
  { key: "equipmentFoundations", group: "construction", sourceLabel: "Фундаменты под оборудование", formLabel: "Фундаменты под оборудование" },
  { key: "stairs", group: "construction", sourceLabel: "Лестницы", formLabel: "Лестницы" },
  { key: "suspendedMetal", group: "construction", sourceLabel: "Подвесные металлоконструкции", formLabel: "Подвесные МК" },
  { key: "blindArea", group: "construction", sourceLabel: "Отмостка, прилегающая территория", formLabel: "Отмостка" },
  { key: "ventilation", group: "construction", sourceLabel: "Система вентиляции", formLabel: "Вентиляция" },
  { key: "utilities", group: "construction", sourceLabel: "Инженерные коммуникации", formLabel: "Инженерные коммуникации" }
];
const A3_TARGETS = [
  { key: "foundations", rowLabel: "Конструкция фундаментов" },
  { key: "outerWalls", rowLabel: "Конструкция несущих стен" },
  { key: "innerWalls", rowLabel: "Конструкция внутренних стен и перегородок" },
  { key: "openings", rowLabel: "Конструкция ворот, дверей и окон" },
  { key: "covering", rowLabel: "Конструкция покрытия" },
  { key: "floorsDrainage", rowLabel: "Тип покрытия пола" },
  // полный текст; разрез с дренажем — бэклог
  { key: "equipmentFoundations", rowLabel: "Конструкция фундаментов под оборудование" },
  { key: "stairs", rowLabel: "Конструкция лестниц" },
  { key: "ventilation", rowLabel: "Конструкция системы вентиляции" },
  { key: "blindArea", rowLabel: "Благоустройство площадки, тип отмостки" }
];
const emptyValues = () => {
  const v = {};
  for (const f of FIELD_DEFS) v[f.key] = "";
  return v;
};
const normalize = (s) => s.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
const cellText = (cellXml) => {
  const parts = [];
  const re = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let m;
  while ((m = re.exec(cellXml)) !== null) parts.push(m[1]);
  return normalize(parts.join(""));
};
const escapeXml = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const setCellPlainText = (cellXml, text) => {
  const safe = escapeXml(text);
  let first = true;
  let out = cellXml.replace(/<w:t([^>]*)>([^<]*)<\/w:t>/g, (_all, attrs) => {
    if (first) {
      first = false;
      const a = /xml:space=/.test(attrs) ? attrs : `${attrs} xml:space="preserve"`;
      return `<w:t${a}>${safe}</w:t>`;
    }
    return `<w:t${attrs}></w:t>`;
  });
  if (first) {
    out = out.replace(
      /<\/w:tcPr>/,
      `</w:tcPr><w:p><w:r><w:t xml:space="preserve">${safe}</w:t></w:r></w:p>`
    );
    if (!out.includes(safe)) {
      out = out.replace(
        /<w:tc([^>]*)>/,
        `<w:tc$1><w:p><w:r><w:t xml:space="preserve">${safe}</w:t></w:r></w:p>`
      );
    }
  }
  return out;
};
const listTables = (documentXml) => {
  const result = [];
  const re = /<w:tbl[\s>][\s\S]*?<\/w:tbl>/g;
  let m;
  while ((m = re.exec(documentXml)) !== null) {
    result.push({ start: m.index, end: m.index + m[0].length, xml: m[0] });
  }
  return result;
};
const listParagraphsBefore = (documentXml, before) => {
  const slice = documentXml.slice(0, before);
  const texts = [];
  const re = /<w:p[\s>][\s\S]*?<\/w:p>/g;
  let m;
  while ((m = re.exec(slice)) !== null) {
    const t = cellText(m[0]);
    if (t) texts.push(t);
  }
  return texts;
};
const headingNear = (documentXml, tableStart) => {
  const before = listParagraphsBefore(documentXml, tableStart);
  return before[before.length - 1] ?? "";
};
const isGeneralTable = (heading, firstRowLabel) => {
  const h = normalize(heading).toLowerCase();
  const row = normalize(firstRowLabel).toLowerCase();
  return row.includes("год постройки") && (h.includes("общие сведения") || h.includes("таблица 5") || h.includes("условия эксплуатации"));
};
const isConstructionTable = (heading, firstRowLabel) => {
  const h = normalize(heading).toLowerCase();
  const row = normalize(firstRowLabel).toLowerCase();
  return row.includes("объемно-планировочные") && (h.includes("объемно-планировочные") || h.includes("таблица 6") || h.includes("конструктивные решения"));
};
const isA3BlockHeading = (heading) => {
  const h = normalize(heading).toLowerCase();
  return h.includes("результаты обследования");
};
const rowLabelAndValueCellIndex = (cells) => {
  if (cells.length >= 3) {
    return { label: cells[1], labelIdx: 1, valueIdx: 2 };
  }
  if (cells.length === 2) {
    return { label: cells[0], labelIdx: 0, valueIdx: 1 };
  }
  return null;
};
const matchLabel = (cellLabel, needle) => {
  const a = normalize(cellLabel).toLowerCase();
  const b = normalize(needle).toLowerCase();
  return a.includes(b) || b.includes(a);
};
const parseTableRowsXml = (tableXml) => {
  const rows = [];
  const trRe = /<w:tr[\s>][\s\S]*?<\/w:tr>/g;
  let tr;
  while ((tr = trRe.exec(tableXml)) !== null) {
    const cellsXml = [];
    const tcRe = /<w:tc[\s>][\s\S]*?<\/w:tc>/g;
    let tc;
    while ((tc = tcRe.exec(tr[0])) !== null) cellsXml.push(tc[0]);
    const texts = cellsXml.map(cellText);
    const meta = rowLabelAndValueCellIndex(texts);
    if (!meta) continue;
    rows.push({ cellsXml, label: meta.label, valueIdx: meta.valueIdx, labelIdx: meta.labelIdx });
  }
  return rows;
};
const readValuesFromTable = (tableXml) => {
  const values = {};
  const rows = parseTableRowsXml(tableXml);
  for (const def of FIELD_DEFS) {
    const row = rows.find((r) => matchLabel(r.label, def.sourceLabel));
    if (!row) continue;
    values[def.key] = cellText(row.cellsXml[row.valueIdx]);
  }
  return values;
};
const writeValuesIntoTable = (tableXml, values, keys) => {
  const allowed = new Set(keys ?? Object.keys(values));
  let xml = tableXml;
  xml = tableXml.replace(/<w:tr[\s>][\s\S]*?<\/w:tr>/g, (trXml) => {
    const cellsXml = [];
    const tcRe = /<w:tc[\s>][\s\S]*?<\/w:tc>/g;
    let tc;
    while ((tc = tcRe.exec(trXml)) !== null) cellsXml.push(tc[0]);
    const texts = cellsXml.map(cellText);
    const meta = rowLabelAndValueCellIndex(texts);
    if (!meta) return trXml;
    const def = FIELD_DEFS.find((d) => matchLabel(meta.label, d.sourceLabel));
    const a3 = A3_TARGETS.find((t) => matchLabel(meta.label, t.rowLabel));
    let key = def?.key ?? a3?.key;
    if (!key || !allowed.has(key)) return trXml;
    const next = values[key];
    if (typeof next !== "string") return trXml;
    const newCells = cellsXml.map(
      (c, idx) => idx === meta.valueIdx ? setCellPlainText(c, next) : c
    );
    let i = 0;
    return trXml.replace(/<w:tc[\s>][\s\S]*?<\/w:tc>/g, () => newCells[i++]);
  });
  return xml;
};
const loadAutofillValues = (templatePath) => {
  const buf = fs.readFileSync(templatePath);
  const zip = new PizZip(buf);
  const file = zip.file("word/document.xml");
  if (!file) throw new Error("В docx нет word/document.xml");
  const documentXml = file.asText();
  const tables = listTables(documentXml);
  const values = emptyValues();
  let foundGeneral = 0;
  let foundConstruction = 0;
  for (const t of tables) {
    const rows = parseTableRowsXml(t.xml);
    if (rows.length === 0) continue;
    const heading = headingNear(documentXml, t.start);
    const firstLabel = rows[0].label;
    if (isGeneralTable(heading, firstLabel) || foundGeneral === 0 && matchLabel(firstLabel, "Год постройки")) {
      if (foundGeneral === 0) Object.assign(values, readValuesFromTable(t.xml));
      foundGeneral++;
      continue;
    }
    if (isConstructionTable(heading, firstLabel) || foundConstruction === 0 && matchLabel(firstLabel, "Объемно-планировочные")) {
      if (foundConstruction === 0) Object.assign(values, readValuesFromTable(t.xml));
      foundConstruction++;
    }
  }
  return { values, foundGeneral, foundConstruction };
};
const fillTablesInDocx = (options) => {
  const { templatePath, outputPath, values } = options;
  const buf = fs.readFileSync(templatePath);
  const zip = new PizZip(buf);
  const file = zip.file("word/document.xml");
  if (!file) throw new Error("В docx нет word/document.xml");
  let documentXml = file.asText();
  const tables = listTables(documentXml);
  let updatedCells = 0;
  const missingTargets = [];
  const hitA3 = /* @__PURE__ */ new Set();
  for (let i = tables.length - 1; i >= 0; i--) {
    const t = tables[i];
    const rows = parseTableRowsXml(t.xml);
    if (rows.length === 0) continue;
    const heading = headingNear(documentXml, t.start);
    const firstLabel = rows[0].label;
    let nextXml = t.xml;
    let changed = false;
    const general = isGeneralTable(heading, firstLabel) || matchLabel(firstLabel, "Год постройки");
    const construction = isConstructionTable(heading, firstLabel) || matchLabel(firstLabel, "Объемно-планировочные");
    const a3 = isA3BlockHeading(heading);
    if (general) {
      const before = nextXml;
      nextXml = writeValuesIntoTable(nextXml, values, FIELD_DEFS.filter((d) => d.group === "general").map((d) => d.key));
      if (nextXml !== before) {
        changed = true;
        updatedCells += FIELD_DEFS.filter((d) => d.group === "general").length;
      }
    } else if (construction) {
      const before = nextXml;
      nextXml = writeValuesIntoTable(nextXml, values, FIELD_DEFS.filter((d) => d.group === "construction").map((d) => d.key));
      if (nextXml !== before) {
        changed = true;
        updatedCells += FIELD_DEFS.filter((d) => d.group === "construction").length;
      }
    } else if (a3) {
      const keys = A3_TARGETS.map((x) => x.key);
      const before = nextXml;
      nextXml = writeValuesIntoTable(nextXml, values, keys);
      if (nextXml !== before) {
        changed = true;
        for (const target of A3_TARGETS) {
          if (rows.some((r) => matchLabel(r.label, target.rowLabel))) {
            hitA3.add(target.rowLabel);
            updatedCells++;
          }
        }
      }
    }
    if (changed) {
      documentXml = documentXml.slice(0, t.start) + nextXml + documentXml.slice(t.end);
    }
  }
  for (const t of A3_TARGETS) {
    if (!hitA3.has(t.rowLabel)) missingTargets.push(t.rowLabel);
  }
  zip.file("word/document.xml", documentXml);
  const out = zip.generate({ type: "nodebuffer" });
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, out);
  return { outputPath, updatedCells, missingTargets };
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
function buildAutofillOutputPath(templatePath) {
  const { dir, name, ext } = path.parse(templatePath);
  return path.join(dir, `${name}_заполненный${ext}`);
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
  electron.ipcMain.handle("autofill:load", async (_e, payload) => {
    try {
      const result = loadAutofillValues(payload.templatePath);
      return { ok: true, ...result };
    } catch (error) {
      return { ok: false, error: humanizeError(error) };
    }
  });
  electron.ipcMain.handle(
    "autofill:run",
    async (_e, payload) => {
      try {
        const outputPath = buildAutofillOutputPath(payload.templatePath);
        const result = fillTablesInDocx({
          templatePath: payload.templatePath,
          outputPath,
          values: payload.values
        });
        return { ok: true, ...result };
      } catch (error) {
        return { ok: false, error: humanizeError(error) };
      }
    }
  );
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
