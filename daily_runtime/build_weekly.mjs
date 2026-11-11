import fs from 'node:fs/promises';
import { Workbook,SpreadsheetFile } from '@oai/artifact-tool';
import {applyCPR,addLegend,verifyCPR,verifyTradingView} from './cpr.mjs';
const job=JSON.parse(await fs.readFile(process.argv[2],'utf8')),p=job.payload;
const w=Workbook.create();
const output=w.worksheets.add('Weekly OHLC'),daily=w.worksheets.add('Daily Prices'),sources=w.worksheets.add('Sources and Method');
output.getRange('A1:E1').values=[['Stock Symbol','Weekly Open','Weekly High','Weekly Low','Weekly Close']];
output.getRange('A2').write(p.weekly.map(r=>[r[0]]));
output.getRange(`B2:E${p.weekly.length+1}`).formulas=p.weekly.map(r=>r.slice(1));
output.getRange('N1').values=[[`Week: ${p.start} to ${p.end} | INR per share`]];
output.getRange('N2').values=[[`${p.count} F&O stocks | Unadjusted cash-market prices`]];
output.getRange('N:N').format.columnWidth=70;
daily.getRange('A1:G1').values=[['Stock Symbol','Date','Open','High','Low','Close','F&O eligible that day']];
daily.getRange('A2').write(p.daily.map(r=>[r[0],new Date(r[1]+'T00:00:00Z'),...r.slice(2)]));
sources.getRange('A1').write(p.notes.map(r=>[...r,...Array(4-r.length).fill(null)]));
sources.getRange('A1').values=[['Report period']];
for(const s of [output,daily,sources]){
 s.showGridLines=false;s.getUsedRange().format.font={name:'Arial',size:11};
 s.getUsedRange().format.rowHeight=21;s.getUsedRange().format.columnWidth=19;
 s.getRange('A:A').format.columnWidth=24;s.freezePanes.freezeRows(1);
}
output.getRange('A1:E1').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'},horizontalAlignment:'center'};
output.getRange(`B2:E${p.weekly.length+1}`).setNumberFormat('#,##0.00');
applyCPR(output,p.weekly.length+1,'WeeklyPrices');
daily.getRange('A1:G1').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'}};
daily.getRange(`B2:B${p.daily.length+1}`).setNumberFormat('dd-mmm-yyyy');
daily.getRange(`C2:F${p.daily.length+1}`).setNumberFormat('#,##0.00');
daily.getRange('G:G').format.columnWidth=26;
sources.getRange('B:B').format.columnWidth=115;sources.getRange('C:C').format.columnWidth=115;sources.getRange('D:D').format.columnWidth=75;
sources.getRange('B1:B11').format.wrapText=true;sources.getRange('A1:D11').format.rowHeight=34;
sources.getRange('A13:D13').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'}};
sources.getRange('A11').values=[['Fetched at (UTC)']];sources.getRange('B11').values=[[new Date(p.notes[10][1])]];
sources.getRange('B11').setNumberFormat('dd-mmm-yyyy hh:mm');
addLegend(w,'weekly');w.recalculate();
const got=output.getRange(`A2:E${p.weekly.length+1}`).values;
for(let i=0;i<got.length;i++)for(let j=0;j<5;j++){
 const expected=p.calculated[i][j];
 if(typeof expected==='number'?Math.abs(got[i][j]-expected)>1e-8:got[i][j]!==expected)throw Error(`Weekly OHLC mismatch ${i},${j}`);
}
verifyCPR(output,p.weekly.length+1);
verifyTradingView(output,p.weekly.length+1);
await (await SpreadsheetFile.exportXlsx(w)).save(job.output);
console.log(`Verified ${got.length} weekly OHLC rows and CPR calculations`);
