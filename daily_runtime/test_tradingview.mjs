import {Workbook} from '@oai/artifact-tool';
import {addTradingViewColumn,verifyTradingView} from './cpr.mjs';
const w=Workbook.create(),s=w.worksheets.add('Symbols');
s.getRange('A1:E1').values=[['Stock Symbol','Open','High','Low','Close']];
const symbols=['ASTRAL','DMART','HINDZINC','M&M','BAJAJ-AUTO','360ONE'];
s.getRange('A2').write(symbols.map(x=>[x]));
addTradingViewColumn(s,symbols.length+1,'SymbolExport');w.recalculate();
verifyTradingView(s,symbols.length+1);
if(s.getRange('L2').formulas[0][0]!=='=A2&","')throw Error('Wrong formula');
s.getRange('A2').values=[['UPDATED']];w.recalculate();
if(s.getRange('L2').values[0][0]!=='UPDATED,')throw Error('Input change not reflected');
// Reorder the symbol inputs to exercise the row references after sorting.
s.getRange('A2').write([...symbols].sort().reverse().map(x=>[x]));w.recalculate();
verifyTradingView(s,symbols.length+1);
if(s.tables.items[0].getHeaderRowRange().values[0].at(-1)!=='TradingView Symbol')throw Error('Export column outside table');
console.log('PASS: exact comma format, special-character symbols, live recalculation and reordered rows; final table column is TradingView Symbol.');
