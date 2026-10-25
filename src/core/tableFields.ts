export type FieldKey =
// таблица 5
  | "yearBuilt"
  | "buildingArea"
  | "buildingVolume"
  | "floors"
  | "fireCategory"
  | "responsibilityLevel"
  | "snowRegion"
  | "windRegion"
  | "seismicity"
  | "climateRegion"
  | "designTemp"
  // таблица 6
  | "volumePlanning"
  | "structuralScheme"
  | "foundations"
  | "outerWalls"
  | "innerWalls"
  | "columns"
  | "covering"
  | "floorSlab"
  | "roof"
  | "openings"
  | "floorsDrainage"
  | "equipmentFoundations"
  | "stairs"
  | "suspendedMetal"
  | "blindArea"
  | "ventilation"
  | "utilities";

export type FieldGroup = "general" | "construction";

export interface FieldDef {
  key: FieldKey;
  group: FieldGroup;
  /** Текст левой ячейки в табл. 5/6 и А.2 (нормализуем при поиске) */
  sourceLabel: string;
  formLabel: string;
}

/** Подписи строк в табл. 5 и 6 — как в шаблоне, можно чуть укоротить для UI */
export const FIELD_DEFS: FieldDef[] = [
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
  { key: "utilities", group: "construction", sourceLabel: "Инженерные коммуникации", formLabel: "Инженерные коммуникации" },
];

export type ValuesMap = Partial<Record<FieldKey, string>>;

/** Где ещё писать значение, кроме sourceLabel в табл. 5/6 и А.2 */
export interface ExtraTarget {
  key: FieldKey;
  /** Подпись строки в таблице А.3 (средняя ячейка) */
  rowLabel: string;
}

/** Только «Конструкция…» / аналоги. Отделка — в бэклоге. */
export const A3_TARGETS: ExtraTarget[] = [
  { key: "foundations", rowLabel: "Конструкция фундаментов" },
  { key: "outerWalls", rowLabel: "Конструкция несущих стен" },
  { key: "innerWalls", rowLabel: "Конструкция внутренних стен и перегородок" },
  { key: "openings", rowLabel: "Конструкция ворот, дверей и окон" },
  { key: "covering", rowLabel: "Конструкция покрытия" },
  { key: "floorsDrainage", rowLabel: "Тип покрытия пола" }, // полный текст; разрез с дренажем — бэклог
  { key: "equipmentFoundations", rowLabel: "Конструкция фундаментов под оборудование" },
  { key: "stairs", rowLabel: "Конструкция лестниц" },
  { key: "ventilation", rowLabel: "Конструкция системы вентиляции" },
  { key: "blindArea", rowLabel: "Благоустройство площадки, тип отмостки" },
];

export const emptyValues = (): ValuesMap => {
  const v: ValuesMap = {};
  for (const f of FIELD_DEFS) v[f.key] = "";
  return v;
};