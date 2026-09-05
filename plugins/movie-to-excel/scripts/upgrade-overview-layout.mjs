import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const defaultTemplate = fileURLToPath(new URL("../assets/movie-archive-template.xlsx", import.meta.url));
const inputPath = path.resolve(process.argv[2] || defaultTemplate);
const outputPath = path.resolve(process.argv[3] || inputPath);

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = wb.worksheets.getItem("我的电影");
const navy = "#183153";
const greenAlt = "#F3FAF0";
const line = "#D7E8D0";

const cards = [
  { columns: ["A", "B", "C", "D"], headerRow: 6, endRow: 11 },
  { columns: ["E", "F", "G", "H"], headerRow: 6, endRow: 11 },
  { columns: ["I", "J", "K", "L"], headerRow: 6, endRow: 11 },
  { columns: ["A", "B", "C", "D"], headerRow: 15, endRow: 20 },
  { columns: ["E", "F", "G", "H"], headerRow: 15, endRow: 20 },
  { columns: ["I", "J", "K", "L"], headerRow: 15, endRow: 20 },
];

for (const { columns, headerRow, endRow } of cards) {
  const first = columns[0];
  const last = columns.at(-1);
  sheet.unmergeCells(`${first}${headerRow}:${last}${endRow}`);
  sheet.getRange(`${first}${headerRow}:${last}${endRow}`).clear({ applyTo: "contents" });
  sheet.getRange(`${first}${headerRow}:${last}${headerRow}`).values = [["片名（中文 / English）", "导演", "年份", "观影日期"]];
  sheet.getRange(`${first}${headerRow}:${last}${headerRow}`).format = {
    fill: navy,
    font: { bold: true, color: "#FFFFFF", fontSize: 9 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    rowHeight: 28,
    borders: { insideVertical: { style: "thin", color: "#FFFFFF" } },
  };
  sheet.getRange(`${first}${headerRow + 1}:${last}${endRow}`).format = {
    fill: greenAlt,
    rowHeight: 23,
    borders: { insideHorizontal: { style: "thin", color: line } },
  };
  sheet.getRange(`${columns[2]}${headerRow + 1}:${columns[2]}${endRow}`).format.numberFormat = "0";
  sheet.getRange(`${columns[3]}${headerRow + 1}:${columns[3]}${endRow}`).format.numberFormat = "yyyy-mm-dd";
}

for (const column of ["A", "E", "I"]) sheet.getRange(`${column}:${column}`).format.columnWidth = 30;
for (const column of ["B", "F", "J"]) sheet.getRange(`${column}:${column}`).format.columnWidth = 20;
for (const column of ["C", "G", "K"]) sheet.getRange(`${column}:${column}`).format.columnWidth = 9;
for (const column of ["D", "H", "L"]) sheet.getRange(`${column}:${column}`).format.columnWidth = 14;

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(JSON.stringify({ ok: true, output: outputPath }, null, 2));
