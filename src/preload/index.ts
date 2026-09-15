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
});