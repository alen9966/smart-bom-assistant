import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const source = process.argv[2];
const outputDir = process.argv[3] || ".qa-subcontract-template";
await fs.mkdir(outputDir, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
console.log((await workbook.inspect({ kind: "workbook,sheet,table", maxChars: 8000, tableMaxRows: 15, tableMaxCols: 18, tableMaxCellChars: 100 })).ndjson);
for (const sheet of workbook.worksheets.items) {
  const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", scale: 1.5, format: "png" });
  await fs.writeFile(`${outputDir}/${sheet.name.replace(/[<>:"/\\|?*]/g, "_")}.png`, new Uint8Array(await preview.arrayBuffer()));
}
