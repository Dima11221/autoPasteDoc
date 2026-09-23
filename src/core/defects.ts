import fs from "node:fs";
import PizZip from "pizzip";

const NO_DEFECTS = "Дефекты и повреждения конструкций не обнаружены";
const HEADING = "Ведомость дефектов и повреждений строительных конструкций";

export const parsePhotoNumbers = (cell: string): number[] => {
  const normalized = cell.replace(/\u00a0/g, " ").trim();

  if (
    !normalized ||
    normalized === "-" ||
    normalized === "–" ||
    normalized === "—"
  ) {
    return [];
  }

  const tokens = normalized
    .split(/[,;]/)
    .map((t) => t.trim())
    .filter(Boolean);

  const result: number[] = [];

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

const cellTextFromXml = (cellXml: string): string => {
  const parts: string[] = [];
  const re = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(cellXml)) !== null) {
    parts.push(match[1]);
  }
  return parts.join("").replace(/\u00a0/g, " ").trim();
};

const extractTables = (documentXml: string): string[][][] => {
  const tables: string[][][] = [];
  const tblRe = /<w:tbl[\s>][\s\S]*?<\/w:tbl>/g;
  let tblMatch: RegExpExecArray | null;

  while ((tblMatch = tblRe.exec(documentXml)) !== null) {
    const tblXml = tblMatch[0];
    const rows: string[][] = [];
    const trRe = /<w:tr[\s>][\s\S]*?<\/w:tr>/g;
    let trMatch: RegExpExecArray | null;

    while ((trMatch = trRe.exec(tblXml)) !== null) {
      const trXml = trMatch[0];
      const cells: string[] = [];
      const tcRe = /<w:tc[\s>][\s\S]*?<\/w:tc>/g;
      let tcMatch: RegExpExecArray | null;

      while ((tcMatch = tcRe.exec(trXml)) !== null) {
        cells.push(cellTextFromXml(tcMatch[0]));
      }
      if (cells.length > 0) rows.push(cells);
    }
    if (rows.length > 0) tables.push(rows);
  }
  return tables;
};

const findDescAndPhotoCols = (header: string[]): { descCol: number; photoCol: number } | null => {
  let descCol = -1;
  let photoCol = -1;

  for (let i = 0; i < header.length; i++) {
    const h = header[i].toLowerCase();
    if (descCol < 0 && h.includes("описание дефектов")) descCol = i;
    if (photoCol < 0 && h.includes("№ фото")) photoCol = i;
  }

  if (descCol >= 0 && photoCol >= 0) {
    return {descCol, photoCol}
  };

  // fallback: столбцы 3 и 4 (1-based) → индексы 2 и 3
  if (header.length >= 4) {
    return { descCol: 2, photoCol: 3 }
  }
  return null;
};

const pickDefectsTable = (tables: string[][][]): string[][] | null => {
  const matches: string[][][] = [];

  for (const table of tables) {
    const header = table[0] ?? [];
    const joined = header.join(" | ");
    if (
      joined.includes("№ фото") &&
      joined.toLowerCase().includes("описание дефектов")
    ) {
      matches.push(table);
    }
  }
  if (matches.length > 0) {
    return matches[matches.length - 1];
  }

  // запасной путь: таблица сразу после заголовка ведомости
  return null;
};



export type CaptionsByOrder = Map<number, string[]>;

export const loadCaptionsFromDocx = (templatePath: string): CaptionsByOrder => {
  const content = fs.readFileSync(templatePath);
  const zip = new PizZip(content);
  const file = zip.file("word/document.xml");
  if (!file) {
    throw new Error("В docx нет word/document.xml");
  };

  const documentXml = file.asText();
  const tables = extractTables(documentXml);
  let table = pickDefectsTable(tables);

  // если по заголовкам не нашли — ищем последнюю таблицу после текста HEADING
  if (!table) {
    const headingPos = documentXml.lastIndexOf(HEADING);
    if (headingPos >= 0) {
      const after = documentXml.slice(headingPos);
      const afterTables = extractTables(after);
      table = afterTables[0] ?? null;
    }
  }

  const captions: CaptionsByOrder = new Map();
  if (!table || table.length < 2) return captions;

  const cols = findDescAndPhotoCols(table[0]);
  if (!cols) return captions;

  const { descCol, photoCol } = cols;
  for(let r = 1; r < table.length; r++) {
    const row = table[r];
    const description = (row[descCol] ?? "").replace(/\s+/g, " ").trim();
    const photoCell = row[photoCol] ?? "";

    if (!description) continue;
    if (description === NO_DEFECTS) continue;

    const numbers = parsePhotoNumbers(photoCell);
    if (numbers.length === 0) continue;

    for (const n of numbers) {
      const list = captions.get(n) ?? [];
      // не дублируем одно и то же описание
      if (!list.includes(description)) list.push(description);
      captions.set(n, list);
    }
  }

  return captions;
};