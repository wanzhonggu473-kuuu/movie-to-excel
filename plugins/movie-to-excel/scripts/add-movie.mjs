import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const CATEGORIES = ["中国大陆", "香港", "台湾", "美国", "欧洲/大洋洲", "其他地区", "待确认"];
const REQUIRED = ["title", "director", "year", "country", "region", "watchedDate", "sourceUrl"];

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (!argv[i].startsWith("--")) continue;
    const key = argv[i].slice(2);
    const next = argv[i + 1];
    args[key] = next && !next.startsWith("--") ? argv[++i] : true;
  }
  return args;
}

function normalize(value) {
  return String(value ?? "").trim().toLocaleLowerCase().replace(/[\s·:：'“”".,，。!！?？-]/g, "");
}

function excelDate(isoDate) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) throw new Error("watchedDate 必须使用 yyyy-mm-dd 格式");
  const date = new Date(`${isoDate}T00:00:00`);
  if (Number.isNaN(date.valueOf())) throw new Error("watchedDate 不是有效日期");
  return date;
}

function validateMovie(movie) {
  const missing = REQUIRED.filter((field) => movie[field] === undefined || String(movie[field]).trim() === "");
  if (missing.length) throw new Error(`缺少必要字段：${missing.join(", ")}`);
  const year = Number(movie.year);
  if (!Number.isInteger(year) || year < 1888 || year > 2200) throw new Error("year 不是合理的上映年份");
  if (!CATEGORIES.includes(movie.region)) throw new Error(`region 必须是：${CATEGORIES.join("、")}`);
  if (movie.rating !== undefined && movie.rating !== "") {
    const rating = Number(movie.rating);
    if (!Number.isFinite(rating) || rating < 0 || rating > 10) throw new Error("rating 必须在 0–10 之间");
  }
}

async function loadMovie(args) {
  if (args["movie-file"]) return JSON.parse(await fs.readFile(path.resolve(args["movie-file"]), "utf8"));
  if (args["movie-json"]) return JSON.parse(args["movie-json"]);
  throw new Error("请通过 --movie-file 或 --movie-json 提供电影信息");
}

function findFirstBlank(rows) {
  const index = rows.findIndex((row) => !String(row?.[1] ?? "").trim());
  if (index < 0) throw new Error("电影数据库已达到 500 条容量，请先扩展模板");
  return index;
}

const args = parseArgs(process.argv.slice(2));
const defaultTemplate = fileURLToPath(new URL("../assets/movie-archive-template.xlsx", import.meta.url));
const inputPath = path.resolve(args.input || defaultTemplate);
const outputPath = path.resolve(args.output || inputPath);
if (inputPath === outputPath && !args["in-place"]) {
  throw new Error("为避免误覆盖，输出路径与输入路径相同时必须显式添加 --in-place");
}

const movie = await loadMovie(args);
validateMovie(movie);

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const db = wb.worksheets.getItem("电影数据库");
const dbRange = db.getRange("A5:L504");
const rows = dbRange.values;

const titleKey = normalize(movie.originalTitle || movie.title);
const duplicate = rows.find((row) => {
  const existingKey = normalize(row?.[2] || row?.[1]);
  return existingKey && existingKey === titleKey && Number(row?.[4]) === Number(movie.year);
});
if (duplicate && !args["allow-duplicate"]) {
  throw new Error(`检测到重复记录：${duplicate[1] || duplicate[2]} (${duplicate[4]})；如确需重复记录，请添加 --allow-duplicate`);
}

const rowIndex = findFirstBlank(rows);
const sheetRow = rowIndex + 5;
const usedCount = rows.filter((row) => String(row?.[1] ?? "").trim()).length;
const recordId = `M${String(usedCount + 1).padStart(4, "0")}`;
const record = [
  recordId,
  String(movie.title).trim(),
  String(movie.originalTitle || "").trim(),
  String(movie.director).trim(),
  Number(movie.year),
  String(movie.country).trim(),
  movie.region,
  String(movie.genre || "").trim(),
  excelDate(movie.watchedDate),
  movie.rating === undefined || movie.rating === "" ? "" : Number(movie.rating),
  String(movie.notes || "").trim(),
  String(movie.sourceUrl).trim(),
];
db.getRange(`A${sheetRow}:L${sheetRow}`).values = [record];
db.getRange(`E${sheetRow}`).format.numberFormat = "0";
db.getRange(`I${sheetRow}`).format.numberFormat = "yyyy-mm-dd";
db.getRange("F:F").format.columnWidth = 28;
db.getRange("G:G").format.columnWidth = 18;

const list = wb.worksheets.getItem("标准清单");
const sourceCols = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"];
list.getRange(`A${sheetRow}:K${sheetRow}`).formulas = [[...sourceCols.map((col) => `=IF('电影数据库'!${col}${sheetRow}="","",'电影数据库'!${col}${sheetRow})`)]];
list.getRange(`D${sheetRow}`).format.numberFormat = "0";
list.getRange(`H${sheetRow}`).format.numberFormat = "yyyy-mm-dd";
list.getRange("E:E").format.columnWidth = 28;
list.getRange("F:F").format.columnWidth = 18;

const updatedRows = db.getRange("A5:L504").values.filter((row) => String(row?.[1] ?? "").trim());
const overview = wb.worksheets.getItem("我的电影");
const topCategories = ["中国大陆", "香港", "台湾"];
const topCount = Math.max(...topCategories.map((category) => updatedRows.filter((row) => row[6] === category).length), 5);
const bottomStartRow = 7 + topCount + 5;
overview.getRange("A13:L15").copyTo(overview.getRange(`A${bottomStartRow - 3}:L${bottomStartRow - 1}`), "all");
overview.getRange(`A${bottomStartRow - 3}:L${bottomStartRow - 1}`).format.rowHeight = 24;
overview.getRange("A13:L504").clear({ applyTo: "contents" });
const cards = {
  "中国大陆": { columns: ["A", "B", "C", "D"], startRow: 7, endRow: 11 },
  "香港": { columns: ["E", "F", "G", "H"], startRow: 7, endRow: 11 },
  "台湾": { columns: ["I", "J", "K", "L"], startRow: 7, endRow: 11 },
  "美国": { columns: ["A", "B", "C", "D"], startRow: bottomStartRow, endRow: bottomStartRow + 4 },
  "欧洲/大洋洲": { columns: ["E", "F", "G", "H"], startRow: bottomStartRow, endRow: bottomStartRow + 4 },
  "其他地区": { columns: ["I", "J", "K", "L"], startRow: bottomStartRow, endRow: bottomStartRow + 4 },
};
for (const [category, card] of Object.entries(cards)) {
  const [titleCol, directorCol, yearCol, dateCol] = card.columns;
  const range = overview.getRange(`${titleCol}${card.startRow}:${dateCol}504`);
  range.clear({ applyTo: "contents" });
  const categoryRows = updatedRows.filter((row) => row[6] === category).slice().reverse();
  categoryRows.forEach((row, index) => {
    const displayTitle = row[2] && normalize(row[2]) !== normalize(row[1]) ? `${row[1]} / ${row[2]}` : row[1];
    const targetRow = card.startRow + index;
    overview.getRange(`${titleCol}${targetRow}:${dateCol}${targetRow}`).values = [[displayTitle, row[3], row[4], row[8]]];
    overview.getRange(`${yearCol}${targetRow}`).format.numberFormat = "0";
    overview.getRange(`${dateCol}${targetRow}`).format.numberFormat = "yyyy-mm-dd";
  });
}

const horizontal = wb.worksheets.getItem("地区横向");
horizontal.getRange("A6:Z504").clear({ applyTo: "contents" });
const horizontalGroups = {
  "中国大陆": { start: 0, country: false }, "香港": { start: 4, country: false },
  "台湾": { start: 8, country: false }, "美国": { start: 12, country: false },
  "欧洲/大洋洲": { start: 16, country: true }, "其他地区": { start: 21, country: true },
};
for (const [category, config] of Object.entries(horizontalGroups)) {
  const categoryRows = updatedRows.filter((row) => row[6] === category);
  categoryRows.forEach((row, index) => {
    const displayTitle = row[2] && normalize(row[2]) !== normalize(row[1]) ? `${row[1]} / ${row[2]}` : row[1];
    const values = [displayTitle, row[3], row[4], row[8]];
    if (config.country) values.push(row[5]);
    horizontal.getRangeByIndexes(5 + index, config.start, 1, values.length).values = [values];
    horizontal.getRangeByIndexes(5 + index, config.start + 2, 1, 1).format.numberFormat = "0";
    horizontal.getRangeByIndexes(5 + index, config.start + 3, 1, 1).format.numberFormat = "yyyy-mm-dd";
  });
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(outputPath);
console.log(JSON.stringify({ ok: true, output: outputPath, recordId, title: movie.title, region: movie.region }, null, 2));
