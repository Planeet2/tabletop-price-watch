import fs from 'node:fs/promises';
import path from 'node:path';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const [input = 'data/latest.csv', output = 'outputs/Tabletop Price Watch.xlsx', preview] = process.argv.slice(2);
const workbook = Workbook.create();
const dashboard = workbook.worksheets.add('Dashboard');
const compare = workbook.worksheets.add('Comparison');
const source = workbook.worksheets.add('Offers');
const stores = ['Miniature Market', 'Cardhaus', 'Noble Knight'];
const navy = '#21354D';
const teal = '#126C64';
const ink = '#203248';
const muted = '#566A7F';
const pale = '#EAF3F4';
const currency = '$#,##0.00';
const write = (sheet, address, values) => { sheet.getRange(address).values = values; };
const formula = (sheet, address, value) => { sheet.getRange(address).formulas = [[value]]; };

const csv = await fs.readFile(input, 'utf8');
const parsed = await Workbook.fromCSV(csv.replace(/^\uFEFF/, ''), { sheetName: 'Input' });
const [headers, ...inputRows] = parsed.worksheets.getItem('Input').getUsedRange().values;
if (headers.length !== 12 || headers[3] !== 'Price USD' || headers[11] !== 'Offer key') {
  throw new Error('CSV does not match the Tabletop Price Watch export schema');
}
const rows = inputRows.filter(row => row[0]).map(row => row.map((value, column) => {
  if (column === 3) {
    if (value === '' || value == null) return null;
    const price = Number(value);
    if (!Number.isFinite(price) || price <= 0) throw new Error(`Invalid price in CSV: ${value}`);
    return price;
  }
  if (column === 9 && value instanceof Date) return value.toISOString().slice(0, 16).replace('T', ' ');
  if (column === 9 && typeof value === 'string') return value.slice(0, 16).replace('T', ' ');
  return value == null ? null : String(value);
}));
const games = [...new Set(rows.map(row => row[0]))];
if (!games.length || rows.length !== games.length * stores.length || rows.some(row => !stores.includes(row[2]))) {
  throw new Error('Expected one record per game and retailer');
}
const keyed = new Map(rows.map(row => [row[11], row]));
if (keyed.size !== rows.length || rows.some(row => row[11] !== `${row[0]}|${row[2]}`)) {
  throw new Error('Offer keys must uniquely identify the game and retailer');
}

function layout(sheet, widths, extent) {
  sheet.showGridLines = false;
  sheet.getRange(extent).format.font.name = 'Arial';
  sheet.getRange(extent).format.font.size = 10;
  sheet.getRange(extent).format.font.color = ink;
  sheet.getRange(extent).format.verticalAlignment = 'center';
  sheet.getRange(extent).format.rowHeight = 20;
  for (const [column, width] of Object.entries(widths)) sheet.getRange(`${column}:${column}`).format.columnWidth = width;
}

function header(sheet, address) {
  const format = sheet.getRange(address).format;
  format.fill = navy;
  format.font.color = '#FFFFFF';
  format.font.bold = true;
  format.horizontalAlignment = 'center';
  format.rowHeight = 26;
}

function title(sheet, address, text) {
  write(sheet, address, [[text]]);
  const format = sheet.getRange(address).format;
  format.font.size = 16;
  format.font.bold = true;
  format.rowHeight = 27;
}

function bands(sheet, first, last, left, right) {
  for (let row = first; row <= last; row += 2) sheet.getRange(`${left}${row}:${right}${row}`).format.fill = '#F3F6F9';
}

const lastOffer = rows.length + 1;
layout(source, { A: 30, B: 15, C: 21, D: 14, E: 18, F: 40, G: 19, H: 17, I: 69, J: 29, K: 36, L: 53 }, `A1:L${lastOffer}`);
write(source, `A1:L${lastOffer}`, [headers, ...rows]);
source.getRange(`D2:D${lastOffer}`).setNumberFormat(currency);
const table = source.tables.add(`A1:L${lastOffer}`, true, 'RetailOffers');
table.style = 'TableStyleMedium2';
header(source, 'A1:L1');
source.freezePanes.freezeRows(1);

const first = 7;
const last = first + games.length - 1;
layout(compare, { A: 33, B: 13, C: 11, D: 18, E: 16, F: 16, G: 13, H: 15, I: 13, J: 15, K: 22, L: 13, M: 13 }, `A1:M${last + 2}`);
compare.tabColor = '#57748E';
title(compare, 'A2', 'Compare by game');
write(compare, 'A4', [['Edit Basket qty to update the basket totals on Dashboard. Blue prices are currently in stock.']]);
compare.getRange('A4').format.font.italic = true;
write(compare, 'A6:M6', [['Game / edition', 'Type', 'Basket qty', 'Miniature Market', 'Status', 'Cardhaus', 'Status', 'Noble Knight', 'Status', 'Best price', 'Lowest offer', 'Gap USD', 'Gap %']]);
header(compare, 'A6:M6');
bands(compare, first, last, 'A', 'M');
compare.getRange(`C${first}:C${last}`).format.fill = '#FFF1CB';
compare.dataValidations.add({ range: `C${first}:C${last}`, rule: { type: 'whole', operator: 'between', formula1: 0, formula2: 20 } });
const basket = new Set(['Cascadia', 'Finspan', 'Azul Duel']);
for (const [index, game] of games.entries()) {
  const row = first + index;
  write(compare, `A${row}:B${row}`, [[game, keyed.get(`${game}|${stores[0]}`)[1]]]);
  write(compare, `C${row}`, [[basket.has(game) ? 1 : 0]]);
  for (const [storeIndex, [price, status]] of [['D','E'],['F','G'],['H','I']].entries()) {
    if (!keyed.has(`${game}|${stores[storeIndex]}`)) throw new Error(`Missing ${game} observation at ${stores[storeIndex]}`);
    const match = `MATCH($A${row}&"|${stores[storeIndex]}",Offers!$L$2:$L$${lastOffer},0)`;
    const offerPrice = `INDEX(Offers!$D$2:$D$${lastOffer},${match})`;
    formula(compare, `${price}${row}`, `=IF(${offerPrice}>0,${offerPrice},"")`);
    formula(compare, `${status}${row}`, `=INDEX(Offers!$E$2:$E$${lastOffer},${match})`);
  }
  const eligible = [['D','E'],['F','G'],['H','I']].map(([p,s]) => `AND(${s}${row}="In stock",ISNUMBER(${p}${row}))`);
  const count = eligible.map(test => `IF(${test},1,0)`).join('+');
  const cheapest = eligible.map((test,i) => `IF(${test},${['D','F','H'][i]}${row},1E+12)`).join(',');
  const highest = eligible.map((test,i) => `IF(${test},${['D','F','H'][i]}${row},0)`).join(',');
  const winners = eligible.map((test,i) => `IF(AND(${test},${['D','F','H'][i]}${row}=J${row}),1,0)`).join('+');
  formula(compare, `J${row}`, `=IF(${count}=0,"",MIN(${cheapest}))`);
  formula(compare, `K${row}`, `=IF(J${row}="","No in-stock offer",IF(${winners}>1,"Tie",IF(AND(${eligible[0]},D${row}=J${row}),"Miniature Market",IF(AND(${eligible[1]},F${row}=J${row}),"Cardhaus","Noble Knight"))))`);
  formula(compare, `L${row}`, `=IF(${count}<2,"",ROUND(MAX(${highest})-J${row},2))`);
  formula(compare, `M${row}`, `=IF(OR(L${row}="",J${row}=0),"",L${row}/J${row})`);
}
for (const column of ['D','F','H','J','L']) compare.getRange(`${column}${first}:${column}${last}`).setNumberFormat(currency);
compare.getRange(`M${first}:M${last}`).setNumberFormat('0.0%');
for (const [p,s] of [['D','E'],['F','G'],['H','I']]) {
  compare.getRange(`${p}${first}:${p}${last}`).conditionalFormats.addCustom(`AND($${s}${first}="In stock",ISNUMBER(${p}${first}))`, { font: { bold: true, color: '#145A8B' } });
}
compare.freezePanes.freezeRows(6);

layout(dashboard, { A: 22, B: 19, C: 13, D: 18, E: 17, F: 14, G: 3, H: 11, I: 11, J: 11, K: 11, L: 11 }, 'A1:L33');
dashboard.tabColor = teal;
title(dashboard, 'A2', 'Tabletop Price Watch');
write(dashboard, 'A3', [['USD · Contiguous US · Published item prices and estimated delivery · Before sales tax']]);
dashboard.getRange('A3').format.font.italic = true;
dashboard.getRange('A3').format.font.color = muted;
write(dashboard, 'A5', [['Run started (UTC)']]);
formula(dashboard, 'B5', '=Offers!J2');
write(dashboard, 'D5', [['Tracked editions']]);
formula(dashboard, 'E5', `=COUNTA(Comparison!A${first}:A${last})`);
write(dashboard, 'G5', [['']]);
write(dashboard, 'H5', [['Comparable at 2+ stores']]);
formula(dashboard, 'K5', `=COUNT(Comparison!L${first}:L${last})`);
dashboard.getRange('A5:L5').format.fill = pale;
dashboard.getRange('A5:L5').format.font.bold = true;
write(dashboard, 'A7', [['Selected basket']]);
dashboard.getRange('A7').format.font.bold = true;
write(dashboard, 'A8:F8', [['Retailer', 'Lines available', 'Subtotal', 'Est. shipping', 'Est. total', 'Above best']]);
header(dashboard, 'A8:F8');
bands(dashboard, 9, 11, 'A', 'F');
write(dashboard, 'A15', [['Shipping assumptions']]);
dashboard.getRange('A15').format.font.bold = true;
write(dashboard, 'A16:E19', [
  ['Retailer', 'Free over/at', 'Base estimate', 'Threshold rule', 'Method'],
  [stores[0], 59, 5.99, 'Over', 'Economy'],
  [stores[1], 99, null, 'At or over', 'Checkout quote'],
  [stores[2], 149, 9.95, 'At or over', 'Economy cap'],
]);
header(dashboard, 'A16:E16');
dashboard.getRange('B17:C19').format.fill = '#FFF1CB';
dashboard.getRange('B17:C19').setNumberFormat(currency);
write(dashboard, 'A13', [['Selected lines']]);
formula(dashboard, 'B13', `=COUNTIF(Comparison!C${first}:C${last},">0")`);
write(dashboard, 'D13', [['Lowest est. basket']]);
formula(dashboard, 'E13', '=IF(COUNT(E9:E11)=0,"",MIN(E9:E11))');
for (const [index, store] of stores.entries()) {
  const row = 9 + index;
  const policy = 17 + index;
  const p = ['D','F','H'][index];
  const s = ['E','G','I'][index];
  const quantities = `Comparison!C${first}:C${last}`;
  const prices = `Comparison!${p}${first}:${p}${last}`;
  const statuses = `Comparison!${s}${first}:${s}${last}`;
  write(dashboard, `A${row}`, [[store]]);
  formula(dashboard, `B${row}`, `=COUNTIFS(${quantities},">0",${statuses},"In stock",${prices},">0")`);
  formula(dashboard, `C${row}`, `=IF(OR(B${row}<>$B$13,$B$13=0),"",ROUND(SUMPRODUCT(${quantities},${prices}),2))`);
  formula(dashboard, `D${row}`, `=IF(C${row}="","",IF(IF(D${policy}="Over",C${row}>B${policy},C${row}>=B${policy}),0,IF(ISNUMBER(C${policy}),C${policy},"Quote")))`);
  formula(dashboard, `E${row}`, `=IF(ISNUMBER(D${row}),ROUND(C${row}+D${row},2),"")`);
  formula(dashboard, `F${row}`, `=IF(ISNUMBER(E${row}),ROUND(E${row}-$E$13,2),"")`);
}
dashboard.getRange('C9:F11').setNumberFormat(currency);
dashboard.getRange('E13').setNumberFormat(currency);
dashboard.getRange('E9:E11').conditionalFormats.addCustom('AND(ISNUMBER(E9),E9=$E$13)', { fill: '#DBF2E9', font: { bold: true, color: '#106353' } });
write(dashboard, 'A21', [['Set quantities on Comparison. A store needs every selected line in stock for a basket estimate.']]);
write(dashboard, 'A22', [['Shipping policies can change. Noble Knight uses its published maximum below the free-shipping threshold.']]);
write(dashboard, 'A23', [['Gap % is the spread divided by the cheapest in-stock price. A single available offer has no spread.']]);
dashboard.getRange('A21:A23').format.font.color = muted;
dashboard.getRange('A21:A23').format.font.italic = true;
write(dashboard, 'A25:E25', [['Retailer', 'Verified listings', 'In-stock offers', 'Lowest price', 'Tied lowest']]);
header(dashboard, 'A25:E25');
bands(dashboard, 26, 28, 'A', 'E');
for (let index = 0; index < stores.length; index++) {
  const row = 26 + index;
  const p = ['D','F','H'][index];
  const s = ['E','G','I'][index];
  const statuses = `Comparison!${s}${first}:${s}${last}`;
  const prices = `Comparison!${p}${first}:${p}${last}`;
  const best = `Comparison!J${first}:J${last}`;
  const winner = `Comparison!K${first}:K${last}`;
  formula(dashboard, `A${row}`, `=A${9 + index}`);
  const verified = ['UPC', 'MPN', 'SKU'].map(method => `COUNTIFS(Offers!C2:C${lastOffer},A${row},Offers!H2:H${lastOffer},"${method}")`);
  formula(dashboard, `B${row}`, `=${verified.join('+')}`);
  formula(dashboard, `C${row}`, `=COUNTIF(${statuses},"In stock")`);
  formula(dashboard, `D${row}`, `=COUNTIF(${winner},A${row})`);
  formula(dashboard, `E${row}`, `=SUMPRODUCT(--(${statuses}="In stock"),--(${prices}=${best}),--(${winner}="Tie"))`);
}
const chart = dashboard.charts.add('bar', [dashboard.getRange('A25:A28'), dashboard.getRange('C25:C28')]);
chart.setPosition('H7', 'L15');
chart.title = 'In-stock offers by retailer';
chart.titleTextStyle.fontSize = 13;
chart.titleTextStyle.typeface = 'Arial';
chart.hasLegend = false;
chart.series.items[0].fill = teal;
chart.xAxis = { textStyle: { typeface: 'Arial', fontSize: 9 } };
chart.yAxis = { numberFormatCode: '0', numberFormatSourceLinked: false, textStyle: { typeface: 'Arial', fontSize: 9 } };

write(dashboard, 'H16', [['Compare one edition']]);
dashboard.getRange('H16').format.font.bold = true;
dashboard.getRange('H17:L17').format.fill = '#FFF1CB';
dashboard.getRange('H17:L17').merge();
write(dashboard, 'H17', [[games.includes('Finspan') ? 'Finspan' : games[0]]]);
dashboard.getRange('H17').dataValidation = { rule: { type: 'list', formula1: `INDIRECT("Comparison!$A$${first}:$A$${last}")` } };
write(dashboard, 'H19:I19', [['Lowest offer', null]]);
dashboard.getRange('H19:I19').merge();
dashboard.getRange('J19:L19').merge();
write(dashboard, 'H20:I20', [['Best price', null]]);
dashboard.getRange('H20:I20').merge();
write(dashboard, 'H21:I21', [['Price gap', null]]);
dashboard.getRange('H21:I21').merge();
const selected = `MATCH(H17,Comparison!A${first}:A${last},0)`;
for (const [cell, column] of [['J19','K'],['J20','J'],['J21','L'],['L21','M']]) {
  formula(dashboard, cell, `=IFERROR(INDEX(Comparison!${column}${first}:${column}${last},${selected}),"")`);
}
dashboard.getRange('J20:J21').setNumberFormat(currency);
dashboard.getRange('L21').setNumberFormat('0.0%');
dashboard.getRange('H19:L21').format.fill = pale;
write(dashboard, 'A30', [['Shipping policy sources']]);
dashboard.getRange('A30').format.font.bold = true;
write(dashboard, 'A31:B33', [
  [stores[0], 'https://www.miniaturemarket.com/utility/customer-service'],
  [stores[1], 'https://www.cardhaus.com/shipping-faq/'],
  [stores[2], 'https://www.nobleknight.com/ordering-and-shipping'],
]);
dashboard.getRange('B31:B33').format.font.color = '#185C8A';

workbook.recalculate();
const originalSelection = dashboard.getRange('H17').values[0][0];
if (games.includes('Cascadia') && games.includes('Finspan') && games.includes('Azul Duel')) {
  const cascadiaRow = first + games.indexOf('Cascadia');
  const originalQty = compare.getRange(`C${cascadiaRow}`).values[0][0];
  const originalCount = dashboard.getRange('B13').values[0][0];
  write(compare, `C${cascadiaRow}`, [[0]]);
  write(dashboard, 'H17', [['Azul Duel']]);
  workbook.recalculate();
  if (dashboard.getRange('B13').values[0][0] !== originalCount - originalQty ||
      dashboard.getRange('J20').values[0][0] !== compare.getRange(`J${first + games.indexOf('Azul Duel')}`).values[0][0]) {
    throw new Error('Dashboard controls did not recalculate');
  }
  write(compare, `C${cascadiaRow}`, [[originalQty]]);
  write(dashboard, 'H17', [[originalSelection]]);
  workbook.recalculate();
}
const keyValues = dashboard.getRange('A8:F13').values;
const compareValues = compare.getRange(`J${first}:M${last}`).values;
console.log(JSON.stringify({ basket: keyValues, comparison: compareValues, coverage: dashboard.getRange('A26:E28').values }));
for (const [sheet, range] of [[dashboard,'A2:L28'],[compare,`A6:M${last}`]]) {
  if (sheet.getRange(range).values.flat().some(value => typeof value === 'string' && /^#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|SPILL!)/.test(value))) {
    throw new Error(`Formula error in ${sheet.name}`);
  }
}
if (preview) {
  for (const [sheet, range, suffix] of [[dashboard,'A1:L33','dashboard'],[compare,`A1:M${last}`,'comparison']]) {
    const png = await workbook.render({ sheetName: sheet.name, range, scale: 1, format: 'png' });
    await fs.writeFile(`${preview}-${suffix}.png`, new Uint8Array(await png.arrayBuffer()));
  }
}
await fs.mkdir(path.dirname(output), { recursive: true });
await (await SpreadsheetFile.exportXlsx(workbook)).save(output);
