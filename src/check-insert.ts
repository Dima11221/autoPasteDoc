import path from "node:path";
import { insertPhotosIntoDocx } from "./core/insert";

async function main() {
  const templatePath = path.resolve(
    "fixtures",
    "960_для тестов.docx",
  );
  const photosFolder = path.resolve("fixtures/photos");
  const outputPath = path.resolve("fixtures/output-test.docx");

  const result = await insertPhotosIntoDocx({
    templatePath,
    photosFolder,
    outputPath,
    // limit: 3,
  });

  console.log(`Вставлено: ${result.insertedCount}`);
  console.log(`Результат: ${result.outputPath}`);
  if (result.skipped.length > 0) {
    console.log("Пропущены:", result.skipped.join(", "));
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});