import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const input = process.argv[2];
const previewDir = process.argv[3];
if (!input) throw new Error("用法：node scripts/verify-workbook.mjs <workbook.xlsx> [preview-dir]");

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path.resolve(input)));
const checks = [
  ["电影数据库", "A1:L8"],
  ["我的电影", "A1:L20"],
  ["地区横向", "A1:Z8"],
  ["标准清单", "A1:K8"],
  ["统计", "A1:B11"],
  ["设置", "A1:B17"],
];
for (const [sheetId, range] of checks) {
  const inspected = await wb.inspect({ kind: "table", sheetId, range, include: "values,formulas", tableMaxRows: 24, tableMaxCols: 26, maxChars: 8000 });
  console.log(inspected.ndjson);
  if (previewDir) {
    await fs.mkdir(previewDir, { recursive: true });
    const image = await wb.render({ sheetName: sheetId, range, scale: 1, format: "png" });
    await fs.writeFile(path.join(previewDir, `${sheetId}.png`), new Uint8Array(await image.arrayBuffer()));
  }
}
const errors = await wb.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 300 }, summary: "formula error scan" });
console.log(errors.ndjson);
