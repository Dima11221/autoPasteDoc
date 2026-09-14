"use strict";
const electron = require("electron");
const path = require("node:path");
const fs = require("node:fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const sharp = require("sharp");
const IMAGE_EXTENSIONS = /* @__PURE__ */ new Set([".jpg", ".jpeg", ".png"]);
const extractOrderNumber = (fileName) => {
  const baseName = path.parse(fileName).name;
  const match = baseName.match(/\s(\d+)$/);
  if (!match) return null;
  return Number(match[1]);
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
  const sizeCache = /* @__PURE__ */ new Map();
  for (const photo of selected) {
    const size = await getPhotoSize(photo.filePath);
    sizeCache.set(photo.filePath, {
      widthCm: size.widthCm,
      heightCm: size.heightCm,
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
  const zip = new PizZip(content);
  const doc = new Docxtemplater(zip, {
    modules: [imageModule],
    paragraphLoop: true,
    linebreaks: true
  });
  doc.render({
    photos: selected.map((photo) => ({ data: photo.filePath }))
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
        const message = error instanceof Error ? error.message : String(error);
        return { ok: false, error: message };
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
