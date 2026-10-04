import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputPath = path.resolve(process.argv[2] || fileURLToPath(new URL("../assets/movie-archive-template.xlsx", import.meta.url)));
const wb = Workbook.create();
const navy = "#1F3A5F";
const cream = "#FFF4C2";
const green = "#F2F8EE";
const border = "#D8E5D2";
const bodyFont = { name: "Arial", size: 10, color: "#1F2937" };
const heading = { fill: navy, font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: false };
const body = { fill: green, font: bodyFont, horizontalAlignment: "left", verticalAlignment: "center", wrapText: false, borders: { preset: "all", style: "thin", color: border } };

function makeSheet(name, title, subtitle, lastColumn) {
  const sheet = wb.worksheets.add(name);
  sheet.showGridLines = false;
  sheet.getRange(`A1:${lastColumn}1`).format = {
    fill: navy, font: { name: "Arial", size: 14, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "left", verticalAlignment: "center", wrapText: false,
  };
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A2").format = { font: { name: "Arial", size: 10, italic: true, color: "#64748B" }, horizontalAlignment: "left", wrapText: false };
  sheet.getRange("1:1").format.rowHeight = 26;
  sheet.getRange("2:2").format.rowHeight = 23;
  return sheet;
}

function setWidths(sheet, widths) {
  for (const [column, width] of Object.entries(widths)) sheet.getRange(`${column}:${column}`).format.columnWidth = width;
}

function addTable(sheet, range, name, headers, widthColumns) {
  const lastColumn = range.split(":")[1].replace(/\d+/g, "");
  const headerRow = Number(range.split(":")[0].replace(/\D+/g, ""));
  sheet.getRange(`${range.split(":")[0]}:${lastColumn}${headerRow}`).values = [headers];
  sheet.getRange(`${range.split(":")[0]}:${lastColumn}${headerRow}`).format = {
    ...heading, borders: { preset: "all", style: "thin", color: "#FFFFFF" },
  };
  sheet.getRangeByIndexes(headerRow, widthColumns[0], 1, headers.length).format = body;
  const table = sheet.tables.add(range, true, name);
  table.style = "TableStyleLight1";
  table.showFilterButton = true;
  return table;
}

const movies = makeSheet("My Movies", "My Movies · Region Overview", "All movies are shown by production region. The second row moves down as the first row grows.", "L");
for (const column of ["A", "E", "I"]) setWidths(movies, { [column]: 37 });
for (const column of ["B", "F", "J"]) setWidths(movies, { [column]: 23 });
for (const column of ["C", "G", "K"]) setWidths(movies, { [column]: 10 });
for (const column of ["D", "H", "L"]) setWidths(movies, { [column]: 16 });
for (const [start, label] of [["A", "Mainland China"], ["E", "Hong Kong"], ["I", "Taiwan"]]) {
  const startIndex = { A: 0, E: 4, I: 8 }[start];
  const cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"].slice(startIndex, startIndex + 4);
  movies.getRange(`${cols[0]}4:${cols[3]}4`).format = heading;
  movies.getRange(`${cols[0]}4`).values = [[label]];
  movies.getRange(`${cols[0]}5:${cols[3]}5`).format = { fill: cream, font: { name: "Arial", size: 10, bold: true, color: navy }, horizontalAlignment: "center" };
  movies.getRange(`${cols[0]}5`).values = [["0 films"]];
  movies.getRange(`${cols[0]}6:${cols[3]}6`).values = [["Title (Chinese / English)", "Director", "Year", "Viewing Date"]];
  movies.getRange(`${cols[0]}6:${cols[3]}6`).format = heading;
  movies.getRange(`${cols[0]}7:${cols[3]}11`).format = body;
}
for (const [start, label] of [["A", "United States"], ["E", "Europe / Oceania"], ["I", "Other Regions"]]) {
  const startIndex = { A: 0, E: 4, I: 8 }[start];
  const cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"].slice(startIndex, startIndex + 4);
  movies.getRange(`${cols[0]}14:${cols[3]}14`).format = heading;
  movies.getRange(`${cols[0]}14`).values = [[label]];
  movies.getRange(`${cols[0]}15:${cols[3]}15`).format = { fill: cream, font: { name: "Arial", size: 10, bold: true, color: navy }, horizontalAlignment: "center" };
  movies.getRange(`${cols[0]}15`).values = [["0 films"]];
  movies.getRange(`${cols[0]}16:${cols[3]}16`).values = [["Title (Chinese / English)", "Director", "Year", "Viewing Date"]];
  movies.getRange(`${cols[0]}16:${cols[3]}16`).format = heading;
  movies.getRange(`${cols[0]}17:${cols[3]}21`).format = body;
}

const region = makeSheet("Region View", "Region View", "Filter directors in every region and countries in the Europe / Oceania and Other Regions groups.", "Z");
const groups = [
  { name: "MainlandChinaFilter", first: "A", last: "D", label: "Mainland China", start: 0, country: false },
  { name: "HongKongFilter", first: "E", last: "H", label: "Hong Kong", start: 4, country: false },
  { name: "TaiwanFilter", first: "I", last: "L", label: "Taiwan", start: 8, country: false },
  { name: "UnitedStatesFilter", first: "M", last: "P", label: "United States", start: 12, country: false },
  { name: "EuropeOceaniaFilter", first: "Q", last: "U", label: "Europe / Oceania", start: 16, country: true },
  { name: "OtherRegionsFilter", first: "V", last: "Z", label: "Other Regions", start: 21, country: true },
];
for (const group of groups) {
  region.getRange(`${group.first}4:${group.last}4`).format = heading;
  region.getRange(`${group.first}4`).values = [[group.label]];
  const headers = ["Title (Chinese / English)", "Director", "Year", "Viewing Date"];
  if (group.country) headers.push("Country / Region");
  addTable(region, `${group.first}5:${group.last}6`, group.name, headers, [group.start]);
  setWidths(region, { [group.first]: 37 });
  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  setWidths(region, { [letters[group.start + 1]]: 23, [letters[group.start + 2]]: 10, [letters[group.start + 3]]: 16 });
  if (group.country) setWidths(region, { [letters[group.start + 4]]: 19 });
}

const list = makeSheet("Master List", "Master List", "The single source of truth for movies. Filter, sort, and add private ratings or notes here.", "K");
const masterHeaders = ["Chinese Title", "Official / Original Title", "Director", "Year", "Country / Region", "Region Group", "Genre", "Viewing Date", "Rating", "Private Notes", "Source URL"];
addTable(list, "A4:K5", "MovieArchiveFilter", masterHeaders, [0]);
setWidths(list, { A: 28, B: 35, C: 24, D: 10, E: 23, F: 20, G: 19, H: 16, I: 10, J: 34, K: 46 });
list.getRange("D5:D500").format.numberFormat = "0";
list.getRange("H5:H500").format.numberFormat = "@";

const stats = makeSheet("Statistics", "Statistics", "Counts are updated locally when movies are added.", "B");
stats.getRange("A3:B3").values = [["Measure", "Count"]];
stats.getRange("A3:B3").format = heading;
stats.getRange("A4:B11").values = [
  ["Total movies", 0], ["Mainland China", 0], ["Hong Kong", 0], ["Taiwan", 0],
  ["United States", 0], ["Europe / Oceania", 0], ["Other Regions", 0], ["Unconfirmed", 0],
];
stats.getRange("A4:B11").format = body;
setWidths(stats, { A: 28, B: 15 });

const settings = makeSheet("Settings", "Settings", "This workbook is stored locally. Public metadata lookup can use the internet.", "B");
settings.getRange("A3:B3").values = [["Preference", "Current Value"]];
settings.getRange("A3:B3").format = heading;
settings.getRange("A4:B8").values = [
  ["Preferred movie view", "My Movies"],
  ["Title order", "Chinese / official English"],
  ["Known full date", "YY-MM-DD"],
  ["Partial date", "Keep the user's precision"],
  ["Internet lookup", "Public metadata only"],
];
settings.getRange("A4:B8").format = body;
setWidths(settings, { A: 29, B: 40 });

const tv = makeSheet("My TV Series", "My TV Series", "A private local record. Add one TV series per row.", "D");
addTable(tv, "A4:D5", "TVSeriesArchive", ["Title (Chinese / English)", "Seasons", "Viewing Period", "Country / Region"], [0]);
setWidths(tv, { A: 42, B: 18, C: 22, D: 22 });
tv.getRange("B5:C500").format.numberFormat = "@";

const docs = makeSheet("My Documentaries", "My Documentaries", "A private local record. Track films and unfinished series in one place.", "F");
addTable(docs, "A4:F5", "DocumentaryArchive", ["Title (Chinese / English)", "Director", "Year", "Seasons / Episodes", "Viewing Period", "Country / Region"], [0]);
setWidths(docs, { A: 42, B: 25, C: 11, D: 29, E: 24, F: 22 });
docs.getRange("C5:C500").format.numberFormat = "0";
docs.getRange("D5:E500").format.numberFormat = "@";

await wb.recalculate();
await fs.mkdir(path.dirname(outputPath), { recursive: true });
await (await SpreadsheetFile.exportXlsx(wb)).save(outputPath);
console.log(JSON.stringify({ output: outputPath, sheets: wb.worksheets.items.map(sheet => sheet.name) }, null, 2));
