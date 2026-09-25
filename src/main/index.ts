import { app, BrowserWindow, dialog, ipcMain, shell } from "electron";
import path from "node:path";
import { insertPhotosIntoDocx } from "../core/insert";
import { humanizeError } from "../core/errors";
import {fillTablesInDocx, loadAutofillValues} from "../core/autofillTables";

function createWindow() {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, "../preload/index.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  if (process.env.ELECTRON_RENDERER_URL) {
    win.loadURL(process.env.ELECTRON_RENDERER_URL);
  } else {
    win.loadFile(path.join(__dirname, "../renderer/index.html"));
  }
}

function buildOutputPath(templatePath: string): string {
  const { dir, name, ext } = path.parse(templatePath);
  return path.join(dir, `${name}_с_фото${ext}`);
}

function buildAutofillOutputPath(templatePath: string): string {
  const { dir, name, ext } = path.parse(templatePath);
  return path.join(dir, `${name}_заполненный${ext}`);
}

function registerIpc() {
  ipcMain.handle("dialog:selectDocx", async () => {
    const result = await dialog.showOpenDialog({
      title: "Выберите Word-документ",
      properties: ["openFile"],
      filters: [{ name: "Word", extensions: ["docx"] }],
    });

    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths[0];
  });

  ipcMain.handle("dialog:selectPhotosFolder", async () => {
    const result = await dialog.showOpenDialog({
      title: "Выберите папку с фото",
      properties: ["openDirectory"],
    });

    if (result.canceled || result.filePaths.length === 0) return null;
    return result.filePaths[0];
  });

  ipcMain.handle(
    "insert:run",
    async (
      _event,
      payload: { templatePath: string; photosFolder: string },
    ) => {
      try {
        const outputPath = buildOutputPath(payload.templatePath);
        const result = await insertPhotosIntoDocx({
          templatePath: payload.templatePath,
          photosFolder: payload.photosFolder,
          outputPath,
        });

        return {
          ok: true as const,
          outputPath: result.outputPath,
          insertedCount: result.insertedCount,
          skipped: result.skipped,
        };
      } catch (error) {

        return { ok: false as const, error: humanizeError(error) };
      }
    },
  );

  ipcMain.handle("shell:showItemInFolder", async (_event, filePath: string) => {
    if (!filePath) return false
    shell.showItemInFolder(filePath);
    return true;
  });

  ipcMain.handle("autofill:load", async (_e, payload: { templatePath: string }) => {
    try {
      const result = loadAutofillValues(payload.templatePath);
      return { ok: true as const, ...result };
    } catch (error) {
      return { ok: false as const, error: humanizeError(error) };
    }
  });

  ipcMain.handle(
    "autofill:run",
    async (
      _e,
      payload: { templatePath: string; values: Record<string, string> },
    ) => {
      try {
        const outputPath = buildAutofillOutputPath(payload.templatePath);
        const result = fillTablesInDocx({
          templatePath: payload.templatePath,
          outputPath,
          values: payload.values,
        });
        return { ok: true as const, ...result };
      } catch (error) {
        return { ok: false as const, error: humanizeError(error) };
      }
    },
  );

}

app.whenReady().then(() => {
  registerIpc();
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});