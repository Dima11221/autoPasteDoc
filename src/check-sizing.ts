import path from "node:path";
import { listPhotosSorted } from "./core/photos";
import {getPhotoSize} from "./core/sizing";

const main = async() => {
  const folder = path.resolve("fixtures/photos");
  const { photos } = listPhotosSorted(folder);

  console.log("order\tор.\tсм(шxв)\tpx\tфайл");

  for (const photo of photos) {
    const size = await getPhotoSize(photo.filePath);
    console.log(
      `${photo.order}\t${size.orientation}\t${size.widthCm.toFixed(2)}x${size.heightCm.toFixed(2)}\t${size.widthPx}x${size.heightPx}\t${photo.fileName}`,
    );
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});