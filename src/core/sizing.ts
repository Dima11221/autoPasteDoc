import sharp from "sharp";

export const LANDSCAPE_WIDTH_CM = 14;
export const PORTRAIT_MAX_CM = 10.5;

export type PhotoSize = {
  widthCm: number;
  heightCm: number;
  widthPx: number;
  heightPx: number;
  orientation: "landscape" | "portrait" | "square";
  orientedBuffer: Buffer;
};

export async function getPhotoSize(filePath: string): Promise<PhotoSize> {
  const { data: orientedBuffer, info } = await sharp(filePath)
    .rotate()
    .jpeg({ quality: 90 })
    .toBuffer({ resolveWithObject: true });

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
      orientedBuffer,
    };
  }

  if (heightPx > widthPx) {
    const heightCm = PORTRAIT_MAX_CM; // 10.5 = максимальная сторона
    const widthCm = heightCm * (widthPx / heightPx);
    return {
      widthCm,
      heightCm,
      widthPx,
      heightPx,
      orientation: "portrait",
      orientedBuffer,
    };
  }

  return {
    widthCm: LANDSCAPE_WIDTH_CM,
    heightCm: LANDSCAPE_WIDTH_CM,
    widthPx,
    heightPx,
    orientation: "square",
    orientedBuffer,
  };
}