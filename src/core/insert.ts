import fs from "node:fs";
import path from "node:path";
import PizZip from "pizzip";
import Docxtemplater, {DXT} from "docxtemplater";

const ImageModule = require("docxtemplater-image-module-free") as new (options: {
  centered: boolean;
  fileType: string;
  getImage: (tagValue: string) => Buffer;
  getSize: (img: Buffer, tagValue: string) => [number, number];
}) => DXT.Module;

import {formatCaption, listPhotosSorted} from "./photos";
import { getPhotoSize } from "./sizing";

function cmToModulePx(cm: number): number {
  const EMU_PER_CM = 360000;
  const EMU_PER_PX = 9525;
  return Math.round((cm * EMU_PER_CM) / EMU_PER_PX);
}

export interface InsertResult {
  outputPath: string;
  insertedCount: number;
  skipped: string[];
};



export async function insertPhotosIntoDocx(options: {
  templatePath: string;
  photosFolder: string;
  outputPath: string;
  limit?: number;
}): Promise<InsertResult> {
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
    const hint =
      skipped.length > 0
        ? `В папке есть изображения без номера в имени (пропущено: ${skipped.length}). Имя должно заканчиваться на пробел и цифру, например "Фасад 1.jpg"`
        : "В папке нет файлов .jpg / .jpeg / .png";
    throw new Error(`Нет фото для вставки. ${hint}`);
  }

  const sizeCache = new Map<
    string,
    {
      widthCm: number;
      heightCm: number;
      orientedBuffer: Buffer;
    }
  >();

  for (const photo of selected) {
    const size = await getPhotoSize(photo.filePath);

    sizeCache.set(photo.filePath, {
      widthCm: size.widthCm,
      heightCm: size.heightCm,
      orientedBuffer: size.orientedBuffer,
    });
  }

  const imageModule = new ImageModule({
    centered: false,
    fileType: "docx",

    getImage: (tagValue: string) => {
      const cached = sizeCache.get(tagValue);
      if (!cached) {
        throw new Error(`Нет данных картинки: ${tagValue}`);
      }
      return cached.orientedBuffer;
    },


    getSize: (_img: Buffer, tagValue: string) => {
      const cached = sizeCache.get(tagValue);
      const widthCm = cached?.widthCm ?? 14;
      const heightCm = cached?.heightCm ?? 14;

      return [cmToModulePx(widthCm), cmToModulePx(heightCm)];
    },
  });

  const content = fs.readFileSync(templatePath);
  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    modules: [imageModule],
    paragraphLoop: true,
    linebreaks: true,
  });

  doc.render({
    photos: selected.map((photo) => ({
      data: photo.filePath,
      caption: formatCaption(photo),
    })),
  });

  const buffer = doc.toBuffer();
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, buffer);

  return {
    outputPath,
    insertedCount: selected.length,
    skipped,
  };
}