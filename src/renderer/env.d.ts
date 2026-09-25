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

type AutofillLoadOk = {
  ok: true;
  values: Record<string, string>;
  foundGeneral: number;
  foundConstruction: number;
};

type AutofillRunOk = {
  ok: true;
  outputPath: string;
  updatedCells: number;
  missingTargets: string[];
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
      showItemInFolder: (filePath: string) => Promise<boolean>;
      loadAutofill: (payload: {
        templatePath: string;
      }) => Promise<AutofillLoadOk | InsertFail>;
      runAutofill: (payload: {
        templatePath: string;
        values: Record<string, string>;
      }) => Promise<AutofillRunOk | InsertFail>;
    };
  }
}