// import sharp from "sharp";
//
// export const WIDTH_LANDSCAPE_CM = 14;
// export const WIDTH_PORTRAIT_CM = 10.5;
//
// export interface PhotoSize {
//   widthCm: number;
//   widthPx: number;
//   heightPx: number;
//   orientation: "landscape" | "portrait" | "square";
// }
//
// export const getPhotoSize = async (filePath: string): Promise<PhotoSize> => {
//   const metadata = await sharp(filePath).metadata();
//
//   const widthPx = metadata.width;
//   const heightPx = metadata.height;
//
//   if (!widthPx || !heightPx) {
//     throw new Error(`Не удалось прочитать: ${filePath}`);
//   }
//
//   if (widthPx > heightPx) {
//     return {
//       widthCm: WIDTH_LANDSCAPE_CM,
//       widthPx,
//       heightPx,
//       orientation: 'landscape',
//     };
//   }
//
//   if (widthPx < heightPx) {
//     return {
//       widthCm: WIDTH_PORTRAIT_CM,
//       widthPx,
//       heightPx,
//       orientation: "portrait",
//     };
//   }
//
//   return {
//     widthCm: WIDTH_LANDSCAPE_CM,
//     widthPx,
//     heightPx,
//     orientation: "square",
//   };
// }


import sharp from "sharp";

export const LANDSCAPE_WIDTH_CM = 14;
export const PORTRAIT_MAX_CM = 10.5;

export type PhotoSize = {
  /** ширина в документе, см */
  widthCm: number;
  /** высота в документе, см */
  heightCm: number;
  widthPx: number;
  heightPx: number;
  orientation: "landscape" | "portrait" | "square";
  /** уже повёрнутые байты картинки (без «кривого» EXIF) */
  orientedBuffer: Buffer;
};

/** Применяем EXIF-поворот и считаем размеры для Word */
export async function getPhotoSize(filePath: string): Promise<PhotoSize> {
  // resolveWithObject даёт info.width/height УЖЕ после rotate()
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