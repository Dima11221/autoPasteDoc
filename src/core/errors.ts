export const humanizeError = (error: unknown): string => {
  const raw = error instanceof Error ? error.message : String(error);
  const lower = raw.toLowerCase();

  //Файл занят / нет прав на запись
  if (
    lower.includes("ebusy") ||
    lower.includes("eperm") ||
    lower.includes("eacces") ||
    lower.includes("resource busy") ||
    lower.includes("operation not permitted")
  ) {
    return "Не удалось сохранить файл. Закройте документ в Word (если открыт) и попробуйте снова.";
  }

  //Нет файла шаблона
  if (lower.includes("enoent")) {
    return "Не удалось найти файл шаблона. Проверьте путь к файлу и попробуйте снова.";
  }

  //Ошибки docxtemplater / плейсхолдера
  if (
    lower.includes("multierror") ||
    lower.includes("unopened_tag") ||
    lower.includes("unclosed_tag") ||
    lower.includes("duplicate_open_tag") ||
    lower.includes("closing_tag_does_not_match") ||
    lower.includes("xmltemplater") ||
    lower.includes("tag \"photos\"") ||
    lower.includes("photos") && lower.includes("not found")
  ) {
    return "В документе не найден корректный блок для фото. Нужны метки \n{#photos}\n{#data}\n\n{/photos}.";
  }

  if (
    lower.includes("zip") ||
    lower.includes("end of central directory")
  ) {
    return "Не удалось прочитать документ. Убедитесь, что выбран именно .docx (не .doc) и файл не поврежден.";
  }

  return raw;
};