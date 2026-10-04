import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
import { displayTitle, insertTableRow, normalize, parseArgs } from "./workbook-utils.mjs";

const args = parseArgs(process.argv.slice(2));
const type = args.type;
if (!['tv', 'documentary'].includes(type)) throw new Error('Use --type tv or --type documentary');
if (!args.input) throw new Error('Pass an existing workbook with --input');
if (!args.output && !args['in-place']) throw new Error('Pass --output <file.xlsx>, or --in-place for an authorized overwrite');
const inputPath = path.resolve(args.input);
const outputPath = path.resolve(args.output || inputPath);
if (inputPath === outputPath && !args['in-place']) throw new Error('Pass --in-place before replacing the input workbook');
const record = args['record-file']
  ? JSON.parse(await fs.readFile(path.resolve(args['record-file']), 'utf8'))
  : args['record-json'] ? JSON.parse(args['record-json']) : null;
if (!record?.title) throw new Error('Pass a title with --record-file or --record-json');

const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = wb.worksheets.getItem(type === 'tv' ? 'My TV Series' : 'My Documentaries');
const tableName = type === 'tv' ? 'TVSeriesArchive' : 'DocumentaryArchive';
const table = sheet.tables.getItem(tableName);
const rows = table.getRange().values.slice(1);
const display = displayTitle(record.title, record.originalTitle);
const titleKey = normalize(record.title);
const matches = rows.map((row, index) => ({ row, index })).filter(({ row }) => {
  const stored = normalize(row?.[0]);
  return stored && (stored === normalize(display) || stored === titleKey || String(row[0]).startsWith(`${record.title} / `));
});
if (matches.length > 1) throw new Error('Multiple matching titles; choose a more specific title');
if (matches.length && !args['update-existing']) throw new Error('This title already exists; use --update-existing to change its progress');

const previous = matches[0]?.row;
const finalTitle = previous && record.originalTitle === undefined ? previous[0] : display;
const value = (name, index) => record[name] === undefined ? previous?.[index] ?? null : record[name] === null ? null : String(record[name]).trim();
let next;
if (type === 'tv') {
  next = [finalTitle, value('seasons', 1), value('viewingPeriod', 2), value('country', 3)];
} else {
  const rawYear = record.year === undefined ? previous?.[2] ?? null : record.year;
  const year = rawYear === null || rawYear === '' ? null : Number(rawYear);
  if (year !== null && (!Number.isInteger(year) || year < 1888 || year > 2200)) throw new Error('Invalid release year');
  next = [finalTitle, value('director', 1), year, value('seasonsOrEpisodes', 3), value('viewingPeriod', 4), value('country', 5)];
}

let rowNumber;
if (matches.length) {
  const header = table.getRange().address.split(':')[0];
  rowNumber = Number(header.replace(/\D/g, '')) + 1 + matches[0].index;
  const last = type === 'tv' ? 'D' : 'F';
  sheet.getRange(`A${rowNumber}:${last}${rowNumber}`).values = [next];
} else {
  rowNumber = insertTableRow(sheet, tableName, next);
}
if (type === 'documentary') sheet.getRange(`C${rowNumber}`).format.numberFormat = '0';
await wb.recalculate();
const errors = await wb.inspect({
  kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
  options: { useRegex: true, maxResults: 100 }, summary: 'final formula error scan',
});
await fs.mkdir(path.dirname(outputPath), { recursive: true });
await (await SpreadsheetFile.exportXlsx(wb)).save(outputPath);
console.log(JSON.stringify({ ok: true, output: outputPath, type, row: rowNumber, record: next, errorScan: errors.ndjson }, null, 2));
