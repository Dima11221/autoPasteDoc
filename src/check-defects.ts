import path from "node:path";
import { loadCaptionsFromDocx } from "./core/defects";

const templatePath = path.resolve("fixtures", "960_для тестов.docx");
const captions = loadCaptionsFromDocx(templatePath);

for (const [order, texts] of [...captions.entries()].sort((a, b) => a[0] - b[0])) {
  console.log(`${order}\t${texts.join(" | ")}`);
}

console.log(`Всего номеров с подписями: ${captions.size}`)