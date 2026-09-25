import { useState, useMemo } from "react";
import {emptyValues, FIELD_DEFS, ValuesMap} from "../core/tableFields";

function fileNameFromPath(fullPath: string): string {
  const parts = fullPath.split(/[/\\]/);
  return parts[parts.length - 1] || fullPath;
}

type Tab = "photos" | "tables";

export const App = () => {
  const [docxPath, setDocxPath] = useState<string | null>(null);
  const [photosFolder, setPhotosFolder] = useState<string | null>(null);
  const [status, setStatus] = useState<string>("Выберите документ и папку с фото.");
  const [busy, setBusy] = useState(false);
  const [lastOutputPath, setLastOutputPath] = useState<string | null>(null);
  const [isError, setIsError] = useState(false);
  const [tab, setTab] = useState<Tab>("photos");
  const [values, setValues] = useState<ValuesMap>(emptyValues);
  const [autofillBusy, setAutofillBusy] = useState(false);


  const docxLabel = useMemo(
    () => (docxPath ? fileNameFromPath(docxPath) : "Файл не выбран"),
    [docxPath],
  );

  const folderLabel = useMemo(
    () => (photosFolder ? fileNameFromPath(photosFolder) : "Папка не выбрана"),
    [photosFolder],
  );

  async function loadTablesFromDocx(templatePath: string) {
    setAutofillBusy(true);
    try {
      const res = await window.api.loadAutofill({ templatePath });
      if (!res.ok) {
        setIsError(true);
        setStatus(`Ошибка чтения таблиц: ${res.error}`);
        return;
      }
      setValues({ ...emptyValues(), ...res.values });
      setIsError(false);
      setStatus(`Таблицы прочитаны (общие: ${res.foundGeneral}, конструкции: ${res.foundConstruction})`);
    } finally {
      setAutofillBusy(false);
    }
  }

  async function onSelectDocx() {
    const selected = await window.api.selectDocx();
    if (!selected) return;
    setDocxPath(selected);
    setLastOutputPath(null);
    setIsError(false);
    setStatus("Документ выбран.");

    if (tab === "tables") {
      await loadTablesFromDocx(selected);
    };
  };

  async function onSelectPhotos() {
    const selected = await window.api.selectPhotosFolder();
    if (!selected) return;
    setPhotosFolder(selected);
    setLastOutputPath(null);
    setIsError(false);
    setStatus("Папка с фото выбрана.");
  }

  async function onInsert() {
    if (!docxPath || !photosFolder) {
      setIsError(true);
      setStatus("Сначала выберите и документ, и папку с фото.");
      return;
    }

    setBusy(true);
    setIsError(false);
    setLastOutputPath(null);
    setStatus("Вставляю фото…");

    try {
      const result = await window.api.insertPhotos({
        templatePath: docxPath,
        photosFolder,
      });

      if (!result.ok) {
        setIsError(true);
        setStatus(`Ошибка: ${result.error}`);
        return;
      }

      const skippedText =
        result.skipped.length > 0
          ? `\nПропущено файлов без номера: ${result.skipped.length}`
          : "";

      setLastOutputPath(result.outputPath);
      setIsError(false);
      setStatus(
        `Готово. Вставлено фото: ${result.insertedCount}\nФайл: ${fileNameFromPath(result.outputPath)}${skippedText}`,
      );
    } catch (error) {
      setIsError(true);
      setStatus(
        error instanceof Error ? error.message : "Неизвестная ошибка",
      );
    } finally {
      setBusy(false);
    }
  }

  const onReveal = async () => {
    if (!lastOutputPath) return;
    await window.api.showItemInFolder(lastOutputPath);
  }

  async function openTablesTab() {
    setTab("tables");
    if (!docxPath) {
      setStatus("Сначала выберите Word-файл на вкладке «Фото».");
      return;
    }
    await loadTablesFromDocx(docxPath);
    setAutofillBusy(true);
    try {
      const res = await window.api.loadAutofill({ templatePath: docxPath });
      if (!res.ok) {
        setIsError(true);
        setStatus(`Ошибка чтения таблиц: ${res.error}`);
        return;
      }
      setValues({ ...emptyValues(), ...res.values });
      setIsError(false);
      setStatus(
        `Таблицы прочитаны (общие: ${res.foundGeneral}, конструкции: ${res.foundConstruction}). Можно править и вставить.`,
      );
    } finally {
      setAutofillBusy(false);
    }
  }

  async function onAutofillInsert() {
    if (!docxPath) return;
    setAutofillBusy(true);
    try {
      const res = await window.api.runAutofill({
        templatePath: docxPath,
        values: values as Record<string, string>,
      });
      if (!res.ok) {
        setIsError(true);
        setStatus(`Ошибка: ${res.error}`);
        return;
      }
      setLastOutputPath(res.outputPath);
      setIsError(false);
      const miss =
        res.missingTargets.length > 0
          ? `\nНе найдено в А.3: ${res.missingTargets.join("; ")}`
          : "";
      setStatus(`Готово. Обновлено ячеек: ~${res.updatedCells}\nФайл: ${fileNameFromPath(res.outputPath)}${miss}`);
    } finally {
      setAutofillBusy(false);
    }
  }

  return (
    <main style={{ fontFamily: "system-ui", padding: 24, maxWidth: 720 }}>
      <nav style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <button disabled={tab === "photos"} onClick={() => setTab("photos")}>Вставка фото</button>
        <button disabled={tab === "tables"} onClick={openTablesTab}>Автозаполнение таблицы</button>
      </nav>

      {tab === "photos" && (
        <div>
          <h1>Вставка фото в Word.</h1>
          <p>Для работы с отчетами. Работает с помощью нумерации в именах фото.</p>

          <div style={{ display: "grid", gap: 12, marginTop: 24 }}>
            <button onClick={onSelectDocx} disabled={busy}>
              1. Выбрать Word-файл (.docx)
            </button>
            <code title={docxPath ?? undefined} style={{whiteSpace: "pre-wrap"}}>
              {docxLabel}
            </code>

            <button onClick={onSelectPhotos} disabled={busy}>
              2. Выбрать папку с фото
            </button>
            <code
              title={photosFolder ?? undefined}
              style={{whiteSpace: "pre-wrap"}}
            >
              {folderLabel}
            </code>

            <button
              onClick={onInsert}
              disabled={busy || !docxPath || !photosFolder}
            >
              {busy ? "Вставляю..." : "3. Вставить фото"}
            </button>

            {busy && (
              <div style={{display: "flex", alignItems: "center", gap: 8}}>
            <span
              aria-hidden
              style={{
                width: 16,
                height: 16,
                border: "2px solid #ccc",
                borderRadius: "50%",
                borderTopColor: "#333",
                display: "inline-block",
                animation: "spin 0.8s linear infinite",
              }}
            />
                <span>Обработка документа, подождите...</span>
              </div>
            )}

            {lastOutputPath && !busy && (
              <button onClick={onReveal}>
                Показать результат в Finder / Проводнике
              </button>
            )}
          </div>

          <pre style={{
            marginTop: 24,
            whiteSpace: "pre-wrap",
            color: isError ? "#b00020" : "inherit",
          }}
          >
        {status}
      </pre>

          <style>{`
        @keyframes spin {
          to {
            transform: rotate(360deg);
          }
        }
      `}</style>
        </div>
      )}

      {tab === "tables" && (
        <div style={{ display: "grid", gap: 12 }}>
          <div style={{ display: "grid", gap: 12, marginTop: 24 }}>
            <button onClick={onSelectDocx} disabled={busy}>
              Выбрать Word-файл (.docx)
            </button>
            <code title={docxPath ?? undefined} style={{whiteSpace: "pre-wrap"}}>
              {docxLabel}
            </code>
          </div>
          <h2>Общие сведения (табл. 5)</h2>
          {FIELD_DEFS.filter((f) => f.group === "general").map((f) => (
            <label key={f.key} style={{ display: "grid", gap: 4 }}>
              <span>{f.formLabel}</span>
              <textarea
                rows={2}
                value={values[f.key] ?? ""}
                onChange={(e) =>
                  setValues((prev) => ({ ...prev, [f.key]: e.target.value }))
                }
              />
            </label>
          ))}
          <h2>Конструкции (табл. 6)</h2>
          {FIELD_DEFS.filter((f) => f.group === "construction").map((f) => (
            <label key={f.key} style={{ display: "grid", gap: 4 }}>
              <span>{f.formLabel}</span>
              <textarea
                rows={3}
                value={values[f.key] ?? ""}
                onChange={(e) =>
                  setValues((prev) => ({ ...prev, [f.key]: e.target.value }))
                }
              />
            </label>
          ))}
          <button
            onClick={onAutofillInsert}
            disabled={autofillBusy || !docxPath}
          >
            {autofillBusy ? "Записываю…" : "Вставить в документ"}
          </button>

          {lastOutputPath && !busy && (
            <button onClick={onReveal}>
              Показать результат в Finder / Проводнике
            </button>
          )}
        </div>
      )}
    </main>
  );
}