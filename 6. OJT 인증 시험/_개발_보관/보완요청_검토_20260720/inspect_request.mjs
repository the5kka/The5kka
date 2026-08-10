import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = String.raw`C:\Users\오국진\Desktop\복사본 OJT 시험 은행 오류 보완 요청List.xlsx`;
const outputDir = path.resolve(".");
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table,drawing",
  maxChars: 20000,
  tableMaxRows: 12,
  tableMaxCols: 12,
  tableMaxCellChars: 200,
});
await fs.writeFile(path.join(outputDir, "summary.ndjson"), summary.ndjson, "utf8");

const sheets = workbook.worksheets.items;
const sheetInfo = [];
for (let index = 0; index < sheets.length; index += 1) {
  const sheet = sheets[index];
  const used = sheet.getUsedRange(false);
  const usedAddress = used ? used.address : null;
  sheetInfo.push({ index, name: sheet.name, usedAddress });

  const details = await workbook.inspect({
    kind: "region,drawing",
    sheetId: sheet.name,
    range: usedAddress ?? undefined,
    maxChars: 30000,
    tableMaxRows: 80,
    tableMaxCols: 20,
    tableMaxCellChars: 500,
  });
  await fs.writeFile(path.join(outputDir, `sheet_${index + 1}_details.ndjson`), details.ndjson, "utf8");

  const preview = await workbook.render({
    sheetName: sheet.name,
    autoCrop: "all",
    scale: 1.6,
    format: "png",
  });
  await fs.writeFile(
    path.join(outputDir, `sheet_${index + 1}.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );

  const focusedPreview = await workbook.render({
    sheetName: sheet.name,
    range: "B2:H8",
    scale: 3,
    format: "png",
  });
  await fs.writeFile(
    path.join(outputDir, `sheet_${index + 1}_focus.png`),
    new Uint8Array(await focusedPreview.arrayBuffer()),
  );
}

await fs.writeFile(path.join(outputDir, "sheets.json"), JSON.stringify(sheetInfo, null, 2), "utf8");
console.log(JSON.stringify(sheetInfo, null, 2));
