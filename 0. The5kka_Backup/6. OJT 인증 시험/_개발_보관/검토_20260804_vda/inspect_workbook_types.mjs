import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = process.argv[2];
if (!inputPath) {
  throw new Error("Usage: node inspect_workbook_types.mjs <workbook>");
}

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const sheetInspection = await workbook.inspect({
  kind: "sheet",
  include: "id,name",
  maxChars: 12000,
});

const types = [
  "\uACF5\uD1B5",
  "\uAC1D\uAD00\uC2DD",
  "\uC8FC\uAD00\uC2DD",
  "VDA",
];

const typeInspection = await workbook.inspect({
  kind: "match",
  searchTerm: `^(?:${types.join("|")})$`,
  options: { useRegex: true, maxResults: 5000 },
  maxChars: 300000,
});

function parseNdjson(text) {
  return text
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => JSON.parse(line));
}

const sheets = parseNdjson(sheetInspection.ndjson)
  .filter((item) => item.kind === "sheet")
  .map((item) => item.name);

const counts = new Map();
const matches = parseNdjson(typeInspection.ndjson)
  .filter(
    (item) =>
      item.kind === "match" &&
      /^E\d+$/i.test(item.address || "") &&
      types.includes(item.value),
  );

for (const item of matches) {
  if (!counts.has(item.sheet)) {
    counts.set(item.sheet, Object.fromEntries(types.map((type) => [type, 0])));
  }
  counts.get(item.sheet)[item.value] += 1;
}

console.log(JSON.stringify({ inputPath, sheets, counts: Object.fromEntries(counts) }, null, 2));
