export {};

type InsertOk = {
  ok: true;
  outputPath: string;
  insertedCount: number;
  skipped: string[];
};

type InsertFail = {
  ok: false;
  error: string;
};

declare global {
  interface Window {
    api: {
      selectDocx: () => Promise<string | null>;
      selectPhotosFolder: () => Promise<string | null>;
      insertPhotos: (payload: {
        templatePath: string;
        photosFolder: string;
      }) => Promise<InsertOk | InsertFail>;
    };
  }
}