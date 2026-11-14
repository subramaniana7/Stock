import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
import {applyCPR,addLegend,verifyCPR,verifyTradingView} from './cpr.mjs';
const job=JSON.parse(await fs.readFile(process.argv[2],'utf8')),p=job.payload;
const w=Workbook.create();
const output=w.worksheets.add('Monthly OHLC'),daily=w.worksheets.add('Daily Prices'),sources=w.worksheets.add('Sources and Method');
output.getRange('A1:E1').values=[['Stock Symbol','Monthly Open','Monthly High','Monthly Low','Monthly Close']];
output.getRange('A2').write(p.rows.map(r=>[r[0]]));
output.getRange(`B2:E${p.rows.length+1}`).formulas=p.rows.map(r=>r.slice(1));
output.getRange('N1').values=[[`Month: ${p.month} | INR per share`]];
output.getRange('N2').values=[[`${p.count} F&O stocks | ${p.sessions.length} trading sessions`]];
daily.getRange('A1:H1').values=[['Stock Symbol','Date','Open','High','Low','Close','F&O eligible that day','Source Cash Symbol']];
daily.getRange('A2').write(p.daily.map(r=>[r[0],new Date(r[1]+'T00:00:00Z'),...r.slice(2)]));
sources.getRange('A1').write(p.notes.map(r=>[...r,...Array(4-r.length).fill(null)]));
for(const s of [output,daily,sources]){
 s.showGridLines=false;s.getUsedRange().format.font={name:'Arial',size:11};
 s.getUsedRange().format.rowHeight=21;s.getUsedRange().format.columnWidth=19;
 s.getRange('A:A').format.columnWidth=24;s.freezePanes.freezeRows(1);
}
output.getRange('N:N').format.columnWidth=70;
output.getRange('A1:E1').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'},horizontalAlignment:'center'};
output.getRange(`B2:E${p.rows.length+1}`).setNumberFormat('#,##0.00');
applyCPR(output,p.rows.length+1,'MonthlyPrices','weekly');
daily.getRange('A1:H1').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'}};
daily.getRange(`B2:B${p.daily.length+1}`).setNumberFormat('dd-mmm-yyyy');
daily.getRange(`C2:F${p.daily.length+1}`).setNumberFormat('#,##0.00');
daily.getRange('G:G').format.columnWidth=26;
daily.getRange('H:H').format.columnWidth=24;
sources.getRange('B:B').format.columnWidth=115;sources.getRange('C:C').format.columnWidth=115;sources.getRange('D:D').format.columnWidth=75;
sources.getRange('B1:B11').format.wrapText=true;sources.getRange('A1:D11').format.rowHeight=42;
sources.getRange('A13:D13').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'}};
sources.getRange('A11').values=[['Fetched at (UTC)']];sources.getRange('B11').values=[[new Date(p.notes[10][1])]];
sources.getRange('B11').setNumberFormat('dd-mmm-yyyy hh:mm');
const guide=addLegend(w,'weekly');
guide.getRange('B9').values=[['Completed calendar-month HLC; levels apply to the following month']];
w.recalculate();
const got=output.getRange(`A2:E${p.rows.length+1}`).values;
for(let i=0;i<got.length;i++)for(let j=0;j<5;j++){
 const expected=p.calculated[i][j];
 if(typeof expected==='number'?Math.abs(got[i][j]-expected)>1e-8:got[i][j]!==expected)throw Error(`Monthly OHLC mismatch ${i},${j}`);
}
verifyCPR(output,p.rows.length+1,'weekly');verifyTradingView(output,p.rows.length+1);
if(job.render_dir){
 await fs.mkdir(job.render_dir,{recursive:true});
 for(const [s,range] of [[output,'A1:L12'],[daily,'A1:H12'],[sources,'A1:B11'],[guide,'A1:C7']]){
  const im=await w.render({sheetName:s.name,range,scale:1.1,format:'png'});
  await fs.writeFile(path.join(job.render_dir,s.name.replaceAll(' ','_')+'.png'),new Uint8Array(await im.arrayBuffer()));
 }
}
await(await SpreadsheetFile.exportXlsx(w)).save(job.output);
console.log(`Verified ${got.length} monthly OHLC rows, CPR bands and TradingView symbols`);
