import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const scripts = path.dirname(fileURLToPath(import.meta.url));
const project = path.resolve(scripts, '../../..');
const outputs = path.join(project, 'outputs');
await fs.mkdir(outputs, {recursive:true});
const scratch = await fs.mkdtemp(path.join(outputs, '.smoke-'));
if (!path.resolve(scratch).startsWith(`${path.resolve(outputs)}${path.sep}`)) throw new Error('Scratch path is outside outputs');
const template = path.resolve(scripts, '../assets/movie-archive-template.xlsx');
const run = (name, args) => execFileSync(process.execPath, [path.join(scripts,name), ...args], {cwd:project,encoding:'utf8'});

try {
  let current = template;
  for (let number = 1; number <= 7; number += 1) {
    const output = path.join(scratch,`movie-${number}.xlsx`);
    const record = {
      title:`测试电影${number}`, originalTitle:`Sample Film ${number}`, director:'测试导演',
      year:2000+number, country:'中国台湾', region:'台湾', watchedDate:number === 1 ? '26' : '2026-09-05',
      sourceUrl:'https://example.org/test-only',
    };
    run('add-movie.mjs',['--input',current,'--output',output,'--movie-json',JSON.stringify(record)]);
    current = output;
  }
  const us = path.join(scratch,'us.xlsx');
  run('add-movie.mjs',['--input',current,'--output',us,'--movie-json',JSON.stringify({
    title:'示例美国电影',originalTitle:'Sample US Film',director:'Sample Director',year:2020,
    country:'United States',region:'美国',watchedDate:null,sourceUrl:'https://example.org/test-only',
  })]);
  let wb = await SpreadsheetFile.importXlsx(await FileBlob.load(us));
  assert.equal(wb.worksheets.getItem('My Movies').getRange('A16').values[0][0],'United States');
  assert.equal(wb.worksheets.getItem('My Movies').getRange('I13').values[0][0],'测试电影7 / Sample Film 7');
  assert.equal(wb.worksheets.getItem('My Movies').getRange('A19').values[0][0],'示例美国电影 / Sample US Film');
  assert.equal(wb.worksheets.getItem('Statistics').getRange('B4').values[0][0],8);
  assert.equal(wb.worksheets.getItem('Region View').tables.getItem('TaiwanFilter').getRange().address,'I5:L12');
  assert.equal(wb.worksheets.getItem('Master List').getRange('H5').values[0][0],'26');
  assert.equal(wb.worksheets.getItem('Master List').getRange('H12').values[0][0],null);

  const tvPath = path.join(scratch,'tv.xlsx');
  run('add-series.mjs',['--type','tv','--input',us,'--output',tvPath,'--record-json',JSON.stringify({
    title:'示例剧集',originalTitle:'Example Series',seasons:'S1',viewingPeriod:'2609',country:'United Kingdom',
  })]);
  const documentaryPath = path.join(scratch,'documentary.xlsx');
  run('add-series.mjs',['--type','documentary','--input',tvPath,'--output',documentaryPath,'--record-json',JSON.stringify({
    title:'示例纪录片',originalTitle:'Example Documentary',year:2015,
    seasonsOrEpisodes:'E1-E3 (in progress)',viewingPeriod:'26-09-28—26-09-30',country:'United States',
  })]);
  const updatedPath = path.join(scratch,'updated.xlsx');
  run('add-series.mjs',['--type','documentary','--input',documentaryPath,'--output',updatedPath,'--update-existing','--record-json',JSON.stringify({
    title:'示例纪录片',seasonsOrEpisodes:'E1-E4 (in progress)',viewingPeriod:'26-09-28—26-10-04',
  })]);
  wb = await SpreadsheetFile.importXlsx(await FileBlob.load(updatedPath));
  assert.deepEqual(wb.worksheets.getItem('My TV Series').getRange('A5:D5').values[0],['示例剧集 / Example Series','S1','2609','United Kingdom']);
  assert.deepEqual(wb.worksheets.getItem('My Documentaries').getRange('A5:F5').values[0],[
    '示例纪录片 / Example Documentary',null,2015,'E1-E4 (in progress)','26-09-28—26-10-04','United States',
  ]);
  assert.equal(wb.worksheets.getItem('Statistics').getRange('B4').values[0][0],8);
  console.log(JSON.stringify({ok:true,movies:8,tv:1,documentaries:1,lowerHeaderRow:16}));
} finally {
  await fs.rm(scratch,{recursive:true,force:true});
}
