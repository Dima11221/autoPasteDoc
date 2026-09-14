import { useState } from "react";

export const App = () => {
  const [docxPath, setDocxPath] = useState<string | null>(null);
  const [photosFolder, setPhotosFolder] = useState<string | null>(null);
  const [status, setStatus] = useState<string>("Выберите документ и папку с фото.");
  const [busy, setBusy] = useState(false);

  async function onSelectDocx() {
    const selected = await window.api.selectDocx();
    if (selected) {
      setDocxPath(selected);
      setStatus("Документ выбран.");
    }
  }

  async function onSelectPhotos() {
    const selected = await window.api.selectPhotosFolder();
    if (selected) {
      setPhotosFolder(selected);
      setStatus("Папка с фото выбрана.");
    }
  }

  async function onInsert() {
    if (!docxPath || !photosFolder) {
      setStatus("Сначала выберите и документ, и папку с фото.");
      return;
    }

    setBusy(true);
    setStatus("Вставляю фото…");

    const result = await window.api.insertPhotos({
      templatePath: docxPath,
      photosFolder,
    });

    setBusy(false);

    if (!result.ok) {
      setStatus(`Ошибка: ${result.error}`);
      return;
    }

    const skippedText =
      result.skipped.length > 0
        ? `\nПропущено файлов: ${result.skipped.length}`
        : "";

    setStatus(
      `Готово. Вставлено: ${result.insertedCount}\nСохранено: ${result.outputPath}${skippedText}`,
    );
  }

  return (
    <main style={{ fontFamily: "system-ui", padding: 24, maxWidth: 720 }}>
      <h1>autoPasteDoc</h1>
      <p>Вставка фото в Word по номерам в именах файлов.</p>

      <div style={{ display: "grid", gap: 12, marginTop: 24 }}>
        <button onClick={onSelectDocx} disabled={busy}>
          1. Выбрать Word-файл (.docx)
        </button>
        <code style={{ whiteSpace: "pre-wrap" }}>
          {docxPath ?? "файл не выбран"}
        </code>

        <button onClick={onSelectPhotos} disabled={busy}>
          2. Выбрать папку с фото
        </button>
        <code style={{ whiteSpace: "pre-wrap" }}>
          {photosFolder ?? "папка не выбрана"}
        </code>

        <button onClick={onInsert} disabled={busy || !docxPath || !photosFolder}>
          3. Вставить фото
        </button>
      </div>

      <pre style={{ marginTop: 24, whiteSpace: "pre-wrap" }}>{status}</pre>
    </main>
  );
}