import fs from "node:fs";
import path from "node:path";
import PizZip from "pizzip";
import {
  A3_TARGETS,
  FIELD_DEFS,
  FieldKey,
  ValuesMap,
  emptyValues,
} from "./tableFields";

const normalize = (s: string): string =>
  s.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();

const cellText = (cellXml: string): string => {
  const parts: string[] = [];
  const re = /<w:t[^>]*>([^<]*)<\/w:t>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(cellXml)) !== null) parts.push(m[1]);
  return normalize(parts.join(""));
};

const escapeXml = (s: string): string =>
  s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

/** Заменить весь текст ячейки одним run — стили ячейки/абзаца сохранятся частично */
const setCellPlainText = (cellXml: string, text: string): string => {
  const safe = escapeXml(text);
  // если есть w:t — заменить содержимое первого, остальные очистить
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
    // ячейка без текста — вставить минимальный paragraph/run
    out = out.replace(
      /<\/w:tcPr>/,
      `</w:tcPr><w:p><w:r><w:t xml:space="preserve">${safe}</w:t></w:r></w:p>`,
    );
    if (!out.includes(safe)) {
      out = out.replace(
        /<w:tc([^>]*)>/,
        `<w:tc$1><w:p><w:r><w:t xml:space="preserve">${safe}</w:t></w:r></w:p>`,
      );
    }
  }
  return out;
};

interface TableHit {
  start: number;
  end: number;
  xml: string;
}

const listTables = (documentXml: string): TableHit[] => {
  const result: TableHit[] = [];
  const re = /<w:tbl[\s>][\s\S]*?<\/w:tbl>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(documentXml)) !== null) {
    result.push({ start: m.index, end: m.index + m[0].length, xml: m[0] });
  }
  return result;
};

const listParagraphsBefore = (documentXml: string, before: number): string[] => {
  const slice = documentXml.slice(0, before);
  const texts: string[] = [];
  const re = /<w:p[\s>][\s\S]*?<\/w:p>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(slice)) !== null) {
    const t = cellText(m[0]); // тот же сбор w:t
    if (t) texts.push(t);
  }
  return texts;
};

const headingNear = (documentXml: string, tableStart: number): string => {
  const before = listParagraphsBefore(documentXml, tableStart);
  return before[before.length - 1] ?? "";
};

const isGeneralTable = (heading: string, firstRowLabel: string): boolean => {
  const h = normalize(heading).toLowerCase();
  const row = normalize(firstRowLabel).toLowerCase();
  return (
    row.includes("год постройки") &&
    (h.includes("общие сведения") || h.includes("таблица 5") || h.includes("условия эксплуатации"))
  );
};

const isConstructionTable = (heading: string, firstRowLabel: string): boolean => {
  const h = normalize(heading).toLowerCase();
  const row = normalize(firstRowLabel).toLowerCase();
  return (
    row.includes("объемно-планировочные") &&
    (h.includes("объемно-планировочные") || h.includes("таблица 6") || h.includes("конструктивные решения"))
  );
};

const isA3BlockHeading = (heading: string): boolean => {
  const h = normalize(heading).toLowerCase();
  return h.includes("результаты обследования");
};

/** 2-col: label|value; 3-col: ?|label|value */
const rowLabelAndValueCellIndex = (cells: string[]): { label: string; valueIdx: number; labelIdx: number } | null => {
  if (cells.length >= 3) {
    return { label: cells[1], labelIdx: 1, valueIdx: 2 };
  }
  if (cells.length === 2) {
    return { label: cells[0], labelIdx: 0, valueIdx: 1 };
  }
  return null;
};

const matchLabel = (cellLabel: string, needle: string): boolean => {
  const a = normalize(cellLabel).toLowerCase();
  const b = normalize(needle).toLowerCase();
  return a.includes(b) || b.includes(a);
};

const parseTableRowsXml = (tableXml: string): { cellsXml: string[]; label: string; valueIdx: number; labelIdx: number }[] => {
  const rows: { cellsXml: string[]; label: string; valueIdx: number; labelIdx: number }[] = [];
  const trRe = /<w:tr[\s>][\s\S]*?<\/w:tr>/g;
  let tr: RegExpExecArray | null;
  while ((tr = trRe.exec(tableXml)) !== null) {
    const cellsXml: string[] = [];
    const tcRe = /<w:tc[\s>][\s\S]*?<\/w:tc>/g;
    let tc: RegExpExecArray | null;
    while ((tc = tcRe.exec(tr[0])) !== null) cellsXml.push(tc[0]);
    const texts = cellsXml.map(cellText);
    const meta = rowLabelAndValueCellIndex(texts);
    if (!meta) continue;
    rows.push({ cellsXml, label: meta.label, valueIdx: meta.valueIdx, labelIdx: meta.labelIdx });
  }
  return rows;
};

const readValuesFromTable = (tableXml: string): ValuesMap => {
  const values: ValuesMap = {};
  const rows = parseTableRowsXml(tableXml);
  for (const def of FIELD_DEFS) {
    const row = rows.find((r) => matchLabel(r.label, def.sourceLabel));
    if (!row) continue;
    values[def.key] = cellText(row.cellsXml[row.valueIdx]);
  }
  return values;
};

const writeValuesIntoTable = (tableXml: string, values: ValuesMap, keys?: FieldKey[]): string => {
  const allowed = new Set(keys ?? (Object.keys(values) as FieldKey[]));
  let xml = tableXml;

  // пересобираем построчно надёжнее, чем replace по всей таблице
  xml = tableXml.replace(/<w:tr[\s>][\s\S]*?<\/w:tr>/g, (trXml) => {
    const cellsXml: string[] = [];
    const tcRe = /<w:tc[\s>][\s\S]*?<\/w:tc>/g;
    let tc: RegExpExecArray | null;
    while ((tc = tcRe.exec(trXml)) !== null) cellsXml.push(tc[0]);
    const texts = cellsXml.map(cellText);
    const meta = rowLabelAndValueCellIndex(texts);
    if (!meta) return trXml;

    const def = FIELD_DEFS.find((d) => matchLabel(meta.label, d.sourceLabel));
    // для А.3 — отдельные подписи
    const a3 = A3_TARGETS.find((t) => matchLabel(meta.label, t.rowLabel));

    let key: FieldKey | undefined = def?.key ?? a3?.key;
    if (!key || !allowed.has(key)) return trXml;
    const next = values[key];
    if (typeof next !== "string") return trXml;

    const newCells = cellsXml.map((c, idx) =>
      idx === meta.valueIdx ? setCellPlainText(c, next) : c,
    );
    // собрать tr заново: заменить последовательность tc
    let i = 0;
    return trXml.replace(/<w:tc[\s>][\s\S]*?<\/w:tc>/g, () => newCells[i++]);
  });

  return xml;
};

export interface LoadResult {
  values: ValuesMap;
  foundGeneral: number; // сколько таблиц «общие сведения» нашли
  foundConstruction: number;
}

export const loadAutofillValues = (templatePath: string): LoadResult => {
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

    if (isGeneralTable(heading, firstLabel) || (foundGeneral === 0 && matchLabel(firstLabel, "Год постройки"))) {
      // берём первую подходящую как эталон для формы (раздел 6 / табл.5)
      if (foundGeneral === 0) Object.assign(values, readValuesFromTable(t.xml));
      foundGeneral++;
      continue;
    }
    if (isConstructionTable(heading, firstLabel) || (foundConstruction === 0 && matchLabel(firstLabel, "Объемно-планировочные"))) {
      if (foundConstruction === 0) Object.assign(values, readValuesFromTable(t.xml));
      foundConstruction++;
    }
  }

  return { values, foundGeneral, foundConstruction };
};

export interface FillResult {
  outputPath: string;
  updatedCells: number;
  missingTargets: string[];
}

export const fillTablesInDocx = (options: {
  templatePath: string;
  outputPath: string;
  values: ValuesMap;
}): FillResult => {
  const { templatePath, outputPath, values } = options;
  const buf = fs.readFileSync(templatePath);
  const zip = new PizZip(buf);
  const file = zip.file("word/document.xml");
  if (!file) throw new Error("В docx нет word/document.xml");

  let documentXml = file.asText();
  const tables = listTables(documentXml);

  let updatedCells = 0;
  const missingTargets: string[] = [];
  const hitA3 = new Set<string>();

  // идём с конца, чтобы индексы start/end не плыли
  for (let i = tables.length - 1; i >= 0; i--) {
    const t = tables[i];
    const rows = parseTableRowsXml(t.xml);
    if (rows.length === 0) continue;
    const heading = headingNear(documentXml, t.start);
    const firstLabel = rows[0].label;

    let nextXml = t.xml;
    let changed = false;

    const general = isGeneralTable(heading, firstLabel) || matchLabel(firstLabel, "Год постройки");
    const construction =
      isConstructionTable(heading, firstLabel) || matchLabel(firstLabel, "Объемно-планировочные");
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