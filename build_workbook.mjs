import fs from 'node:fs/promises';
import path from 'node:path';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const [input = 'data/latest.csv', output = 'outputs/Tabletop Price Watch.xlsx', preview] = process.argv.slice(2);
const workbook = Workbook.create();
const international = workbook.worksheets.add('Dashboard');
const markets = workbook.worksheets.add('Markets');
const dashboard = workbook.worksheets.add('US Basket');
const compare = workbook.worksheets.add('Comparison');
const source = workbook.worksheets.add('Offers');
const stores = ['Miniature Market', 'Cardhaus', 'Noble Knight'];
const foreignStores = ['Philibert', 'Edinorog'];
const marketNames = ['United States', 'France', 'Russia'];
const marketCurrencies = ['USD', 'EUR', 'RUB'];
const navy = '#21354D';
const teal = '#126C64';
const ink = '#203248';
const muted = '#566A7F';
const pale = '#EAF3F4';
const currency = '$#,##0.00';
const fxNumber = '0.000000000';
const write = (sheet, address, values) => { sheet.getRange(address).values = values; };
const formula = (sheet, address, value) => { sheet.getRange(address).formulas = [[value]]; };

const csv = await fs.readFile(input, 'utf8');
const parsed = await Workbook.fromCSV(csv.replace(/^\uFEFF/, ''), { sheetName: 'Input' });
const [csvHeaders, ...inputRows] = parsed.worksheets.getItem('Input').getUsedRange().values;
const extraHeaders = ['Market', 'Local price', 'Currency', 'USD per unit', 'FX date', 'FX source'];
const legacy = csvHeaders.length === 12;
if (csvHeaders[3] !== 'Price USD' || csvHeaders[11] !== 'Offer key' ||
    (!legacy && csvHeaders.slice(12).join('|') !== extraHeaders.join('|'))) {
  throw new Error('CSV does not match the Tabletop Price Watch export schema');
}
const headers = legacy ? [...csvHeaders, ...extraHeaders] : csvHeaders;
const rows = inputRows.filter(row => row[0]).map(record => {
  const row = legacy ? [...record, marketNames[0], record[3], 'USD', record[3] ? 1 : null, null, null] : record;
  return row.map((value, column) => {
    if ([3, 13, 15].includes(column)) {
      if (value === '' || value == null) return null;
      const price = Number(value);
      if (!Number.isFinite(price) || price <= 0) throw new Error(`Invalid price in CSV: ${value}`);
      return price;
    }
    if ([9, 16].includes(column) && value instanceof Date) {
      return value.toISOString().slice(0, column === 16 ? 10 : 16).replace('T', ' ');
    }
    if (column === 9 && typeof value === 'string') return value.slice(0, 16).replace('T', ' ');
    return value == null ? null : String(value);
  });
});
const games = [...new Set(rows.map(row => row[0]))];
const trackedStores = legacy ? stores : [...stores, ...foreignStores];
if (!games.length || rows.length !== games.length * trackedStores.length || rows.some(row => !trackedStores.includes(row[2]))) {
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

function offerValue(column, game, seller) {
  return `INDEX(Offers!$${column}$2:$${column}$${lastOffer},MATCH(${game}&"|"&${seller},Offers!$L$2:$L$${lastOffer},0))`;
}

layout(source, { A: 30, B: 15, C: 21, D: 14, E: 18, F: 40, G: 19, H: 17, I: 69, J: 29, K: 36, L: 53,
  M: 17, N: 15, O: 12, P: 17, Q: 16, R: 75 }, `A1:R${lastOffer}`);
write(source, `A1:R${lastOffer}`, [headers, ...rows]);
source.getRange(`D2:D${lastOffer}`).setNumberFormat(currency);
source.getRange(`N2:N${lastOffer}`).setNumberFormat('#,##0.00');
source.getRange(`P2:P${lastOffer}`).setNumberFormat(fxNumber);
const table = source.tables.add(`A1:R${lastOffer}`, true, 'RetailOffers');
table.style = 'TableStyleMedium2';
header(source, 'A1:R1');
source.freezePanes.freezeRows(1);

const first = 7;
const last = first + games.length - 1;
layout(compare, { A: 33, B: 13, C: 11, D: 18, E: 16, F: 16, G: 13, H: 15, I: 13, J: 15, K: 22, L: 13, M: 13 }, `A1:M${last + 2}`);
compare.tabColor = '#57748E';
title(compare, 'A2', 'Compare US retailers by game');
write(compare, 'A4', [['Edit Basket qty to update the totals on US Basket. Blue prices were in stock at collection.']]);
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
    const offerPrice = offerValue('D', `$A${row}`, JSON.stringify(stores[storeIndex]));
    formula(compare, `${price}${row}`, `=IF(${offerPrice}>0,${offerPrice},"")`);
    formula(compare, `${status}${row}`, `=${offerValue('E', `$A${row}`, JSON.stringify(stores[storeIndex]))}`);
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
title(dashboard, 'A2', 'US basket and retailer comparison');
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

layout(markets, { A: 35, B: 15, C: 17, D: 16, E: 16, F: 16, G: 19, H: 16, I: 15, J: 14 }, `A1:J${last}`);
markets.tabColor = '#57748E';
title(markets, 'A2', 'Prices by market');
write(markets, 'A3', [['Lowest in-stock USD price for the same edition in each market.']]);
write(markets, 'A4', [['Spread is the highest market minimum minus the lowest. Percent is relative to the lowest.']]);
markets.getRange('A3:A4').format.font.color = muted;
write(markets, 'A6:J6', [['Game / edition', 'Type', 'United States', 'France', 'Russia', 'Best USD', 'Lowest market',
  'Spread USD', 'Spread %', 'Markets priced']]);
header(markets, 'A6:J6');
bands(markets, first, last, 'A', 'J');
for (let row = first; row <= last; row++) {
  formula(markets, `A${row}`, `=Comparison!A${row}`);
  formula(markets, `B${row}`, `=Comparison!B${row}`);
  formula(markets, `C${row}`, `=IF(ISNUMBER(Comparison!J${row}),Comparison!J${row},"")`);
  for (const [index, seller] of foreignStores.entries()) {
    const store = JSON.stringify(seller);
    const price = offerValue('D', `$A${row}`, store);
    const status = offerValue('E', `$A${row}`, store);
    const identity = offerValue('H', `$A${row}`, store);
    formula(markets, `${['D','E'][index]}${row}`, `=IFERROR(IF(AND(${status}="In stock",ISNUMBER(${price}),${price}>0,${identity}="UPC"),${price},""),"")`);
  }
  formula(markets, `J${row}`, `=COUNT(C${row}:E${row})`);
  formula(markets, `F${row}`, `=IF(J${row}=0,"",MIN(C${row}:E${row}))`);
  formula(markets, `G${row}`, `=IF(J${row}=0,"",IF(J${row}=1,"Only one market",IF(COUNTIF(C${row}:E${row},F${row})>1,"Tie",INDEX($C$6:$E$6,1,MATCH(F${row},C${row}:E${row},0)))))`);
  formula(markets, `H${row}`, `=IF(J${row}<2,"",ROUND(MAX(C${row}:E${row})-F${row},2))`);
  formula(markets, `I${row}`, `=IF(J${row}<2,"",H${row}/F${row})`);
}
markets.getRange(`C${first}:F${last}`).setNumberFormat(currency);
markets.getRange(`H${first}:H${last}`).setNumberFormat(currency);
markets.getRange(`I${first}:I${last}`).setNumberFormat('0.0%');
markets.getRange(`C${first}:E${last}`).conditionalFormats.addCustom(`AND(ISNUMBER(C${first}),C${first}=$F${first},$J${first}>=2)`,
  { fill: '#DBF2E9', font: { bold: true, color: '#106353' } });
markets.freezePanes.freezeRows(6);

layout(international, { A: 17, B: 22, C: 16, D: 20, E: 14, F: 16, G: 3, H: 11, I: 11, J: 11, K: 11, L: 11, M: 11 }, 'A1:M25');
international.tabColor = teal;
title(international, 'A2', 'Tabletop Price Watch');
write(international, 'A3', [['United States · France · Russia  |  Same English-language editions, listed item prices in USD']]);
international.getRange('A3').format.font.color = muted;
write(international, 'A5:F5', [['Run started (UTC)', null, 'Tracked editions', null, 'In 2+ markets', null]]);
formula(international, 'B5', '=Offers!J2');
formula(international, 'D5', `=COUNTA(Markets!A${first}:A${last})`);
formula(international, 'F5', `=COUNTIF(Markets!J${first}:J${last},">=2")`);
international.getRange('A5:M5').format.fill = pale;
international.getRange('A5:M5').format.font.bold = true;
international.getRange('D5:F5').format.horizontalAlignment = 'center';
write(international, 'A7', [['Game / edition']]);
international.getRange('A7').format.font.bold = true;
international.getRange('B7:F7').merge();
international.getRange('B7:F7').format.fill = '#FFF1CB';
const defaultGame = games.includes('Finspan') ? 'Finspan' : games[0];
write(international, 'B7', [[defaultGame]]);
international.getRange('B7').dataValidation = { rule: { type: 'list', formula1: `INDIRECT("Markets!$A$${first}:$A$${last}")` } };
write(international, 'A9:F9', [['Market', 'Retailer', 'Local price', 'Availability', 'Price USD', 'Above best']]);
header(international, 'A9:F9');
bands(international, 10, 12, 'A', 'F');
const globalSelection = `MATCH($B$7,Markets!$A$${first}:$A$${last},0)`;
const marketResult = column => `INDEX(Markets!$${column}$${first}:$${column}$${last},${globalSelection})`;
const usWinner = `INDEX(Comparison!$K$${first}:$K$${last},${globalSelection})`;
for (const [index, market] of marketNames.entries()) {
  const row = 10 + index;
  write(international, `A${row}`, [[market]]);
  const eligiblePrice = marketResult(['C','D','E'][index]);
  if (index === 0) {
    formula(international, 'B10', `=IFERROR(IF(ISNUMBER(${eligiblePrice}),${usWinner},""),"")`);
    formula(international, 'C10', `=IFERROR(IF(ISNUMBER(${eligiblePrice}),${eligiblePrice},""),"")`);
    formula(international, 'D10', '=IF(ISNUMBER(C10),"In stock","No in-stock offer")');
    formula(international, 'E10', '=IF(ISNUMBER(C10),C10,"")');
  } else {
    const seller = foreignStores[index - 1];
    write(international, `B${row}`, [[seller]]);
    for (const [column, offerColumn] of [['C','N'],['E','D']]) {
      const price = offerValue(offerColumn, '$B$7', `B${row}`);
      formula(international, `${column}${row}`, `=IFERROR(IF(${price}>0,${price},""),"")`);
    }
    const status = offerValue('E', '$B$7', `B${row}`);
    formula(international, `D${row}`, `=IFERROR(${status},"No listing")`);
  }
  formula(international, `F${row}`, `=IFERROR(IF(AND(ISNUMBER(${eligiblePrice}),$E$15>=2),ROUND(${eligiblePrice}-$B$14,2),""),"")`);
  international.getRange(`C${row}`).setNumberFormat(`#,##0.00" ${marketCurrencies[index]}"`);
}
international.getRange('E10:F12').setNumberFormat(currency);
international.getRange('C10:C12').format.horizontalAlignment = 'center';
international.getRange('E10:E12').conditionalFormats.addCustom('AND(ISNUMBER(F10),F10=0)', { fill: '#DBF2E9', font: { bold: true, color: '#106353' } });
write(international, 'A14:E15', [['Best USD', null, null, 'Lowest market', null], ['Spread USD', null, null, 'Markets priced', null]]);
for (const [cell, column, numeric] of [['B14','F',true],['E14','G',false],['B15','H',true],['E15','J',true],['C15','I',true]]) {
  const value = marketResult(column);
  formula(international, cell, numeric ? `=IFERROR(IF(ISNUMBER(${value}),${value},""),"")` : `=IFERROR(IF(ISTEXT(${value}),${value},""),"")`);
}
international.getRange('B14:B15').setNumberFormat(currency);
international.getRange('C15').setNumberFormat('0.0%');
international.getRange('B14:C15').format.horizontalAlignment = 'center';
international.getRange('A14:F15').format.fill = pale;
international.getRange('A14:F15').format.font.bold = true;
write(international, 'A17', [['Listed prices before delivery, duties or tax adjustments.']]);
write(international, 'A18', [['Only verified in-stock offers with a USD price qualify.']]);
write(international, 'A19', [['See Markets for all editions and US Basket for domestic delivery estimates.']]);
international.getRange('A17:A19').format.font.color = muted;

write(international, 'A21:D21', [['Currency', 'USD per unit', 'Rate date', 'Source']]);
header(international, 'A21:D21');
bands(international, 22, 23, 'A', 'D');
for (const [index, unit] of marketCurrencies.slice(1).entries()) {
  const row = 22 + index;
  write(international, `A${row}`, [[unit]]);
  const match = `MATCH(A${row},Offers!$O$2:$O$${lastOffer},0)`;
  const rate = `INDEX(Offers!$P$2:$P$${lastOffer},${match})`;
  formula(international, `B${row}`, `=IFERROR(IF(${rate}>0,${rate},""),"")`);
  formula(international, `C${row}`, `=IFERROR(IF(ISNUMBER(B${row}),INDEX(Offers!$Q$2:$Q$${lastOffer},${match}),""),"")`);
  write(international, `D${row}`, [[index === 0 ? 'European Central Bank' : 'Bank of Russia']]);
}
international.getRange('B22:B23').setNumberFormat(fxNumber);
international.getRange('B22:C23').format.horizontalAlignment = 'center';

write(international, 'A27:B27', [['Market', 'In-stock USD']]);
for (const [index, col] of ['C','D','E'].entries()) {
  const row = 28 + index;
  formula(international, `A${row}`, `=A${10 + index}`);
  const value = marketResult(col);
  formula(international, `B${row}`, `=IFERROR(IF(ISNUMBER(${value}),${value},""),"")`);
}
header(international, 'A27:B27');
international.getRange('B28:B30').setNumberFormat(currency);
const marketChart = international.charts.add('bar', [international.getRange('A27:A30'), international.getRange('B27:B30')]);
marketChart.setPosition('H7', 'M19');
marketChart.title = 'Lowest in-stock price by market';
marketChart.titleTextStyle.fontSize = 13;
marketChart.titleTextStyle.typeface = 'Arial';
marketChart.hasLegend = false;
marketChart.series.items[0].fill = teal;
marketChart.xAxis = { textStyle: { typeface: 'Arial', fontSize: 9 } };
marketChart.yAxis = { numberFormatCode: '$0', numberFormatSourceLinked: false, textStyle: { typeface: 'Arial', fontSize: 9 } };

workbook.recalculate();
const originalSelection = dashboard.getRange('H17').values[0][0];
if (games.includes('Cascadia') && games.includes('Finspan') && games.includes('Azul Duel')) {
  const cascadiaRow = first + games.indexOf('Cascadia');
  const originalQty = compare.getRange(`C${cascadiaRow}`).values[0][0];
  const originalCount = dashboard.getRange('B13').values[0][0];
  write(compare, `C${cascadiaRow}`, [[0]]);
  write(dashboard, 'H17', [['Azul Duel']]);
  write(international, 'B7', [['Azul Duel']]);
  workbook.recalculate();
  if (dashboard.getRange('B13').values[0][0] !== originalCount - originalQty ||
      dashboard.getRange('J20').values[0][0] !== compare.getRange(`J${first + games.indexOf('Azul Duel')}`).values[0][0] ||
      international.getRange('B14').values[0][0] !== markets.getRange(`F${first + games.indexOf('Azul Duel')}`).values[0][0]) {
    throw new Error('Dashboard controls did not recalculate');
  }
  write(compare, `C${cascadiaRow}`, [[originalQty]]);
  write(dashboard, 'H17', [[originalSelection]]);
  write(international, 'B7', [[defaultGame]]);
  workbook.recalculate();
}
const keyValues = dashboard.getRange('A8:F13').values;
const compareValues = compare.getRange(`J${first}:M${last}`).values;
console.log(JSON.stringify({ basket: keyValues, comparison: compareValues, markets: markets.getRange(`A7:J${last}`).values,
  selected: international.getRange('A9:F15').values }));
for (const [sheet, range] of [[international,'A2:M30'],[markets,`A6:J${last}`],[dashboard,'A2:L28'],[compare,`A6:M${last}`]]) {
  if (sheet.getRange(range).values.flat().some(value => typeof value === 'string' && /^#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|SPILL!)/.test(value))) {
    throw new Error(`Formula error in ${sheet.name}`);
  }
}
if (preview) {
  for (const [sheet, range, suffix] of [[international,'A1:M24','dashboard'],[markets,`A1:J${last}`,'markets'],
    [dashboard,'A1:L33','us-basket'], [compare,`A1:M${last}`,'comparison'], [source,'M1:R12','offers']]) {
    const png = await workbook.render({ sheetName: sheet.name, range, scale: 1, format: 'png' });
    await fs.writeFile(`${preview}-${suffix}.png`, new Uint8Array(await png.arrayBuffer()));
  }
}
await fs.mkdir(path.dirname(output), { recursive: true });
await (await SpreadsheetFile.exportXlsx(workbook)).save(output);
