export const REGIONS = [
  { key: "中国大陆", label: "Mainland China", first: "A", table: "MainlandChinaFilter", start: 0, country: false },
  { key: "香港", label: "Hong Kong", first: "E", table: "HongKongFilter", start: 4, country: false },
  { key: "台湾", label: "Taiwan", first: "I", table: "TaiwanFilter", start: 8, country: false },
  { key: "美国", label: "United States", first: "A", table: "UnitedStatesFilter", start: 12, country: false },
  { key: "欧洲/大洋洲", label: "Europe / Oceania", first: "E", table: "EuropeOceaniaFilter", start: 16, country: true },
  { key: "其他地区", label: "Other Regions", first: "I", table: "OtherRegionsFilter", start: 21, country: true },
];

export function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    if (!argv[i].startsWith("--")) continue;
    const key = argv[i].slice(2);
    const next = argv[i + 1];
    args[key] = next && !next.startsWith("--") ? argv[++i] : true;
  }
  return args;
}

export function normalize(value) {
  return String(value ?? "").trim().toLocaleLowerCase().replace(/[\s·:：'“”".,，。!！?？/\-–—]/g, "");
}

function validFullDate(year, month, day) {
  const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
  return date.getUTCFullYear() === Number(year) && date.getUTCMonth() + 1 === Number(month) && date.getUTCDate() === Number(day);
}

export function viewingDate(value) {
  if (value === undefined) {
    const now = new Date();
    const yy = String(now.getFullYear()).slice(-2);
    return `${yy}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
  }
  if (value === null || value === "") return "";
  const text = String(value).trim();
  const match = /^(\d{2}|\d{4})(?:-(\d{2})(?:-(\d{2}))?)?$/.exec(text);
  if (!match) throw new Error("Viewing date must be YY, YY-MM, YY-MM-DD, or the corresponding four-digit-year form");
  const [, year, month, day] = match;
  if (month && (Number(month) < 1 || Number(month) > 12)) throw new Error("Invalid viewing month");
  if (day && !validFullDate(year.length === 2 ? `20${year}` : year, month, day)) throw new Error("Invalid viewing date");
  return `${year.slice(-2)}${month ? `-${month}` : ""}${day ? `-${day}` : ""}`;
}

export function displayTitle(chinese, original) {
  const first = String(chinese ?? "").trim();
  const second = String(original ?? "").trim();
  return first && second && normalize(first) !== normalize(second) ? `${first} / ${second}` : first || second;
}

export function masterRows(wb) {
  const table = wb.worksheets.getItem("Master List").tables.getItem("MovieArchiveFilter");
  return table.getRange().values.slice(1).filter(row => String(row?.[0] ?? "").trim());
}

export function insertTableRow(sheet, tableName, record) {
  const table = sheet.tables.getItem(tableName);
  const values = table.getRange().values;
  const headerAddress = table.getRange().address.split(":")[0];
  const firstDataRow = Number(headerAddress.replace(/\D/g, "")) + 1;
  if (values.length === 2 && values[1].every(cell => cell === null || cell === "")) {
    const firstColumn = headerAddress.replace(/\d/g, "");
    const start = firstColumn.charCodeAt(0) - 65;
    sheet.getRangeByIndexes(firstDataRow - 1, start, 1, record.length).values = [record];
    return firstDataRow;
  }
  table.rows.add(null, [record]);
  return firstDataRow + values.length - 1;
}

export function refreshOverview(wb) {
  const sheet = wb.worksheets.getItem("My Movies");
  const rows = masterRows(wb);
  const topMax = Math.max(5, ...REGIONS.slice(0, 3).map(region => rows.filter(row => row[5] === region.key).length));
  const lowerLabelRow = 7 + topMax + 2;
  sheet.getRange("A4:L1000").clear({ applyTo: "all" });
  const navy = "#1F3A5F", cream = "#FFF4C2", green = "#F2F8EE", border = "#D8E5D2";
  const heading = { fill: navy, font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: false };
  const countStyle = { fill: cream, font: { name: "Arial", size: 10, bold: true, color: navy }, horizontalAlignment: "center", verticalAlignment: "center" };
  const bodyStyle = { fill: green, font: { name: "Arial", size: 10, color: "#1F2937" }, horizontalAlignment: "left", verticalAlignment: "center", wrapText: false, borders: { preset: "all", style: "thin", color: border } };
  for (const [index, region] of REGIONS.entries()) {
    const first = region.first;
    const last = String.fromCharCode(first.charCodeAt(0) + 3);
    const labelRow = index < 3 ? 4 : lowerLabelRow;
    const countRow = labelRow + 1, headerRow = labelRow + 2, dataRow = labelRow + 3;
    const records = rows.filter(row => row[5] === region.key);
    sheet.getRange(`${first}${labelRow}:${last}${labelRow}`).format = heading;
    sheet.getRange(`${first}${labelRow}`).values = [[region.label]];
    sheet.getRange(`${first}${countRow}:${last}${countRow}`).format = countStyle;
    sheet.getRange(`${first}${countRow}`).values = [[`${records.length} films`]];
    sheet.getRange(`${first}${headerRow}:${last}${headerRow}`).values = [["Title (Chinese / English)", "Director", "Year", "Viewing Date"]];
    sheet.getRange(`${first}${headerRow}:${last}${headerRow}`).format = heading;
    const reserve = Math.max(5, records.length);
    sheet.getRange(`${first}${dataRow}:${last}${dataRow + reserve - 1}`).format = bodyStyle;
    if (records.length) {
      const rendered = records.map(row => [displayTitle(row[0], row[1]), row[2], row[3], row[7]]);
      sheet.getRange(`${first}${dataRow}:${last}${dataRow + records.length - 1}`).values = rendered;
      sheet.getRangeByIndexes(dataRow - 1, first.charCodeAt(0) - 65 + 2, records.length, 1).format.numberFormat = "0";
      sheet.getRangeByIndexes(dataRow - 1, first.charCodeAt(0) - 65 + 3, records.length, 1).format.numberFormat = "@";
    }
  }
}

export function refreshStatistics(wb) {
  const rows = masterRows(wb);
  const sheet = wb.worksheets.getItem("Statistics");
  sheet.getRange("B4:B11").values = [
    [rows.length], ...REGIONS.map(region => [rows.filter(row => row[5] === region.key).length]),
    [rows.filter(row => row[5] === "待确认").length],
  ];
}
