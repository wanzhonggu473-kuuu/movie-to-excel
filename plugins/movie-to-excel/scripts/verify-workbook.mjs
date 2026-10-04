import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const input = process.argv[2];
const previewDir = process.argv[3];
if (!input) throw new Error('Usage: node scripts/verify-workbook.mjs <workbook.xlsx> [preview-dir]');
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path.resolve(input)));
const expected = ['My Movies', 'Region View', 'Master List', 'Statistics', 'Settings', 'My TV Series', 'My Documentaries'];
const actual = wb.worksheets.items.map(sheet => sheet.name);
if (JSON.stringify(expected) !== JSON.stringify(actual)) throw new Error(`Unexpected sheets: ${actual.join(', ')}`);
const summary = actual.map(name => {
  const sheet = wb.worksheets.getItem(name);
  return { name, used: sheet.getUsedRange().address, tables: sheet.tables.items.map(table => ({ name: table.name, range: table.getRange().address })) };
});
const master = wb.worksheets.getItem('Master List').tables.getItem('MovieArchiveFilter').getRange().values.slice(1).filter(row => String(row?.[0] ?? '').trim());
const stats = wb.worksheets.getItem('Statistics').getRange('B4').values[0][0];
if (stats !== master.length) throw new Error(`Movie count mismatch: Master List=${master.length}, Statistics=${stats}`);
const errors = await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:100},summary:'formula error scan'});
if (previewDir) {
  await fs.mkdir(path.resolve(previewDir), {recursive:true});
  for (const [name, range] of [['My Movies','A1:L22'],['Region View','A1:Z7'],['My TV Series','A1:D7'],['My Documentaries','A1:F7']]) {
    const image = await wb.render({sheetName:name,range,scale:1,format:'png'});
    await fs.writeFile(path.join(path.resolve(previewDir),`${name.replaceAll(' ','-')}.png`),new Uint8Array(await image.arrayBuffer()));
  }
}
console.log(JSON.stringify({summary,movies:master.length,errorScan:errors.ndjson},null,2));
