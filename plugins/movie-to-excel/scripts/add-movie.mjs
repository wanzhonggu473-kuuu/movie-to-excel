import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
import { REGIONS, displayTitle, insertTableRow, masterRows, normalize, parseArgs, refreshOverview, refreshStatistics, viewingDate } from "./workbook-utils.mjs";

const args = parseArgs(process.argv.slice(2));
const template = fileURLToPath(new URL("../assets/movie-archive-template.xlsx", import.meta.url));
const inputPath = path.resolve(args.input || template);
if (!args.output && !args["in-place"]) throw new Error("Pass --output <file.xlsx>, or --in-place for a previously authorized overwrite");
const outputPath = path.resolve(args.output || inputPath);
if (inputPath === outputPath && !args["in-place"]) throw new Error("Pass --in-place before replacing the input workbook");

const movie = args["movie-file"]
  ? JSON.parse(await fs.readFile(path.resolve(args["movie-file"]), "utf8"))
  : args["movie-json"] ? JSON.parse(args["movie-json"]) : null;
if (!movie) throw new Error("Pass verified metadata with --movie-file or --movie-json");
const required = ["title", "director", "year", "country", "region", "sourceUrl"];
const missing = required.filter(key => !String(movie[key] ?? "").trim());
if (missing.length) throw new Error(`Missing verified movie metadata: ${missing.join(", ")}`);
if (![...REGIONS.map(region => region.key), "待确认"].includes(movie.region)) throw new Error("Unknown region group");
const year = Number(movie.year);
if (!Number.isInteger(year) || year < 1888 || year > 2200) throw new Error("Invalid release year");
const rating = movie.rating === undefined || movie.rating === null || movie.rating === "" ? null : Number(movie.rating);
if (rating !== null && (!Number.isFinite(rating) || rating < 0 || rating > 10)) throw new Error("Rating must be between 0 and 10");

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const existing = masterRows(wb);
const titleKeys = [normalize(movie.title), normalize(movie.originalTitle)].filter(Boolean);
const duplicate = existing.some(row => Number(row[3]) === year && [row[0], row[1]].some(value => titleKeys.includes(normalize(value))));
if (duplicate && !args["allow-duplicate"]) throw new Error("This title and year already exist; use --allow-duplicate only for an intended repeat viewing");

const list = wb.worksheets.getItem("Master List");
const watched = viewingDate(movie.watchedDate);
const record = [
  String(movie.title).trim(), String(movie.originalTitle ?? "").trim(), String(movie.director).trim(),
  year, String(movie.country).trim(), movie.region, String(movie.genre ?? "").trim(),
  watched, rating, String(movie.notes ?? "").trim(), String(movie.sourceUrl).trim(),
];
const listRow = insertTableRow(list, "MovieArchiveFilter", record);
list.getRange(`D${listRow}`).format.numberFormat = "0";
list.getRange(`H${listRow}`).format.numberFormat = "@";

const region = REGIONS.find(item => item.key === movie.region);
if (region) {
  const view = wb.worksheets.getItem("Region View");
  const regionRecord = [displayTitle(movie.title, movie.originalTitle), String(movie.director).trim(), year, watched];
  if (region.country) regionRecord.push(String(movie.country).trim());
  const viewRow = insertTableRow(view, region.table, regionRecord);
  view.getRangeByIndexes(viewRow - 1, region.start + 2, 1, 1).format.numberFormat = "0";
  view.getRangeByIndexes(viewRow - 1, region.start + 3, 1, 1).format.numberFormat = "@";
}
refreshOverview(wb);
refreshStatistics(wb);
await wb.recalculate();
const errors = await wb.inspect({
  kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan",
});
await fs.mkdir(path.dirname(outputPath), { recursive: true });
await (await SpreadsheetFile.exportXlsx(wb)).save(outputPath);
console.log(JSON.stringify({ ok: true, output: outputPath, title: movie.title, region: movie.region, viewed: watched, errorScan: errors.ndjson }, null, 2));
