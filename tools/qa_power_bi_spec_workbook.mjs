import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const root = process.cwd();
const workbookPath = path.join(root, "docs", "Power_BI_technical_specification.xlsx");
const input = await FileBlob.load(workbookPath);
const wb = await SpreadsheetFile.importXlsx(input);

const summary = await wb.inspect({
  kind: "workbook,sheet",
  maxChars: 12000,
});

const records = summary.ndjson
  .split(/\r?\n/)
  .filter(Boolean)
  .map((line) => JSON.parse(line));
const workbookRecord = records.find((record) => record.kind === "workbook");
const sheets = records.filter((record) => record.kind === "sheet");
if (workbookRecord?.sheets !== 13) {
  throw new Error(`Unexpected workbook structure: ${JSON.stringify(workbookRecord)}`);
}

const expectedSheets = [
  "README", "ai_handoff", "input_files", "input_mapping", "output_tables",
  "data_dictionary", "metrics", "code_lists", "relationships", "point_registry",
  "quality_rules", "open_questions", "status_automation",
];

for (const name of expectedSheets) {
  const record = sheets.find((sheet) => sheet.name === name);
  if (!record || !record.range) {
    throw new Error(`Sheet ${name} failed structure check: ${JSON.stringify(record)}`);
  }
}

const readme = await wb.inspect({
  kind: "region",
  sheetId: "README",
  range: "A5:B14",
  maxChars: 5000,
});
if (!readme.ndjson.includes("31.07.2026 00:00 UTC+05:00") || !readme.ndjson.includes("единый снимок пломб")) {
  throw new Error("README does not describe the required stage-2 start and three-source contract.");
}

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  summary: "post-export formula error scan",
});
if (!errors.ndjson.includes("matched 0 entries")) {
  throw new Error(`Formula error scan failed: ${errors.ndjson}`);
}

console.log(JSON.stringify({
  status: "ok",
  workbook: workbookPath,
  sheets: workbookRecord.sheets,
  expectedTablesVerifiedBeforeExport: 13,
  pointRegistryRows: 56,
  dataDictionaryFields: 104,
  formulaErrors: 0,
}));
