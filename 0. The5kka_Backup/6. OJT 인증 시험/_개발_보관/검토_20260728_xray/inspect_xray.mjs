import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "../../작업자 인증 시험 문제_20260727.xlsm";
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const sheets = await workbook.inspect({
  kind: "sheet",
  include: "id,name",
  maxChars: 8000,
});
console.log("SHEETS");
console.log(sheets.ndjson);

const matches = await workbook.inspect({
  kind: "match",
  sheetId: "X-Ray 일반용",
  searchTerm: "이상발생 처리절차",
  options: { maxResults: 20 },
  maxChars: 8000,
});
console.log("MATCHES");
console.log(matches.ndjson);

const region = await workbook.inspect({
  kind: "region",
  sheetId: "X-Ray 일반용",
  range: "B1:Q160",
  maxChars: 30000,
  tableMaxRows: 160,
  tableMaxCols: 16,
  tableMaxCellChars: 300,
});
console.log("REGION");
console.log(region.ndjson);
