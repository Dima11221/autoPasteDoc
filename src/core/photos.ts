import fs from "node:fs";
import path from "node:path";

const IMAGE_EXTENSIONS = new Set([".jpg", ".jpeg", ".png"]);

export type PhotoItem = {
  filePath: string;
  fileName: string;
  order: number;
};

export const extractOrderNumber = (fileName: string): number | null => {
  const baseName = path.parse(fileName).name;
  const match = baseName.match(/\s(\d+)$/);
  if (!match) return null;
  return Number(match[1]);
};

export const formatCaption = (photo: PhotoItem): string => {
  return `Фото ${photo.order}.`;
};

export const listPhotosSorted = (folderPath: string): {
  photos: PhotoItem[];
  skipped: string[];
} => {
  const entries = fs.readdirSync(folderPath);
  const photos: PhotoItem[] = [];
  const skipped: string[] = [];

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
      order,
    });
  }

  photos.sort((a, b) => a.order - b.order);
  return { photos, skipped };
};