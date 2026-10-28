import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = process.argv[2];
if (!inputPath) throw new Error("Usage: node qa_power_bi_output.mjs <workbook.xlsx>");

const expected = {
  "01_Перевозки": ["ID", "Дата_среза", "Этап", "Показатель", "Направление", "Транспорт", "Пункт_пропуска_место", "Код_страны", "Страна", "Количество_перевозок", "Источник"],
  "02_Пломбы": ["ID", "Дата_среза", "Место", "Тип_места", "Остаток_НП", "Использовано_НП_всего", "В_пути", "Использовано_НП_текущий_месяц", "Примечание"],
  "03_Страны": ["Column1", "Column2", "Column3"],
  "04_Места (3)": ["Место", "Широта", "Долгота", "Регион", "Тип", "Комментарий", "Пункт_перевозки"],
  "Количество перевозок": ["Column", "Количество перевозок"],
};

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const names = workbook.worksheets.items.map((sheet) => sheet.name);
if (JSON.stringify(names) !== JSON.stringify(Object.keys(expected))) {
  throw new Error(`Unexpected sheets: ${JSON.stringify(names)}`);
}

const previewDir = path.join(path.dirname(inputPath), "power_bi_preview");
await fs.mkdir(previewDir, { recursive: true });

for (const [sheetName, headers] of Object.entries(expected)) {
  const sheet = workbook.worksheets.getItem(sheetName);
  const actual = sheet.getRangeByIndexes(0, 0, 1, headers.length).values[0];
  if (JSON.stringify(actual) !== JSON.stringify(headers)) {
    throw new Error(`${sheetName}: unexpected headers ${JSON.stringify(actual)}`);
  }
  const preview = await workbook.render({
    sheetName,
    range: `A1:${String.fromCharCode(64 + headers.length)}${sheetName === "01_Перевозки" ? 45 : 30}`,
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(previewDir, `${sheetName.replace(/[<>:"/\\|?*]/g, "_")}.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

console.log(JSON.stringify({ inputPath, sheets: names, previewDir }, null, 2));
