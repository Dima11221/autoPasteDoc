import { contextBridge, ipcRenderer } from "electron";

export type InsertOk = {
  ok: true;
  outputPath: string;
  insertedCount: number;
  skipped: string[];
};

export type InsertFail = {
  ok: false;
  error: string;
};

export type AutofillLoadOk = {
  ok: true;
  values: Record<string, string>;
  foundGeneral: number;
  foundConstruction: number;
};

export type AutofillRunOk = {
  ok: true;
  outputPath: string;
  updatedCells: number;
  missingTargets: string[];
};

contextBridge.exposeInMainWorld("api", {
  selectDocx: (): Promise<string | null> =>
    ipcRenderer.invoke("dialog:selectDocx"),

  selectPhotosFolder: (): Promise<string | null> =>
    ipcRenderer.invoke("dialog:selectPhotosFolder"),

  insertPhotos: (payload: {
    templatePath: string;
    photosFolder: string;
  }): Promise<InsertOk | InsertFail> =>
    ipcRenderer.invoke("insert:run", payload),

  showItemInFolder: (filePath: string): Promise<boolean> =>
    ipcRenderer.invoke("shell:showItemInFolder", filePath),

  loadAutofill: (payload: { templatePath: string }): Promise<AutofillLoadOk | InsertFail> =>
    ipcRenderer.invoke("autofill:load", payload),

  runAutofill: (payload: {
    templatePath: string;
    values: Record<string, string>;
  }): Promise<AutofillRunOk | InsertFail> =>
    ipcRenderer.invoke("autofill:run", payload),
});