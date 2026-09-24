import { useState, useMemo } from "react";

function fileNameFromPath(fullPath: string): string {
  const parts = fullPath.split(/[/\\]/);
  return parts[parts.length - 1] || fullPath;
}

export const App = () => {
  const [docxPath, setDocxPath] = useState<string | null>(null);
  const [photosFolder, setPhotosFolder] = useState<string | null>(null);
  const [status, setStatus] = useState<string>("Выберите документ и папку с фото.");
  const [busy, setBusy] = useState(false);
  const [lastOutputPath, setLastOutputPath] = useState<string | null>(null);
  const [isError, setIsError] = useState(false);

  const docxLabel = useMemo(
    () => (docxPath ? fileNameFromPath(docxPath) : "Файл не выбран"),
    [docxPath],
  );

  const folderLabel = useMemo(
    () => (photosFolder ? fileNameFromPath(photosFolder) : "Папка не выбрана"),
    [photosFolder],
  );

  async function onSelectDocx() {
    const selected = await window.api.selectDocx();
    if (!selected) return;
    setDocxPath(selected);
    setLastOutputPath(null);
    setIsError(false);
    setStatus("Документ выбран.");
  }

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

  return (
    <main style={{ fontFamily: "system-ui", padding: 24, maxWidth: 720 }}>
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
    </main>
  );
}