import path from "node:path";
import { listPhotosSorted } from "./core/photos";

const folder = path.resolve("fixtures/photos");
const { photos, skipped } = listPhotosSorted(folder);

console.log("Будут вставлены:");

for (const photo of photos) {
  console.log(`${photo.order}\t${photo.fileName}`);
}

if (skipped.length > 0) {
  console.log("\nПропущены (нет номера перед расширением):");
  for (const name of skipped) {
    console.log(`- ${name}`);
  }
} else {
  console.log("\nНет пропущенных файлов.");
}

console.log(`Всего к вставке: ${photos.length}`);