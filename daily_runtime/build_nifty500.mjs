import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
import {applyCPR,addLegend,expectedCPR,verifyTradingView} from './cpr.mjs';
const job=JSON.parse(await fs.readFile(process.argv[2],'utf8')),p=job.payload;
const w=Workbook.create(),label=p.kind==='weekly'?'Weekly':'Monthly';
const output=w.worksheets.add('NSE500 '+label),daily=w.worksheets.add('Daily Prices'),coverage=w.worksheets.add('Coverage'),members=w.worksheets.add('Constituents'),sources=w.worksheets.add('Sources and Method');
output.getRange('A1:E1').values=[['Stock Symbol',label+' Open',label+' High',label+' Low',label+' Close']];
output.getRange('A2').write(p.rows.map(r=>[r[0]]));
output.getRange(`B2:E${p.count+1}`).formulas=p.rows.map(r=>r.slice(1).map(x=>x.startsWith('=')?x:'=""'));
output.getRange('N1').values=[[`Nifty 500 current list | ${p.start} to ${p.end} | INR`]];
output.getRange('N2').values=[[`${p.count} index entries; list downloaded ${p.constituents.downloaded_at.slice(0,10)}`]];
daily.getRange('A1:H1').values=[['Stock Symbol','Date','Open','High','Low','Close','Source Cash Symbol','Source Cash Series']];
daily.getRange('A2').write(p.daily.map(r=>[r[0],new Date(r[1]+'T00:00:00Z'),...r.slice(2)]));
coverage.getRange('A1:G1').values=[['Stock Symbol','Coverage status','Available sessions','Expected sessions','First observation','Last observation','Constituent ISIN']];
coverage.getRange('A2').write(p.coverage.map(r=>[...r.slice(0,4),r[4]?new Date(r[4]+'T00:00:00Z'):'',r[5]?new Date(r[5]+'T00:00:00Z'):'',r[6]]));
coverage.getRange(`E2:F${p.count+1}`).setNumberFormat('dd-mmm-yyyy');
members.getRange('A1:E1').values=[['Company Name','Industry','Symbol','Series in index list','ISIN Code']];
members.getRange('A2').write(p.members.map(r=>[r['Company Name'],r.Industry,r.Symbol,r.Series,r['ISIN Code']]));
sources.getRange('A1').write(p.notes.map(r=>[...r,...Array(4-r.length).fill(null)]));
for(const s of [output,daily,coverage,members,sources]){
 s.showGridLines=false;s.getUsedRange().format.font={name:'Arial',size:11};s.getUsedRange().format.rowHeight=21;s.getUsedRange().format.columnWidth=19;
 s.getRange('A:A').format.columnWidth=24;s.freezePanes.freezeRows(1);
}
for(const [s,last] of [[output,'E'],[daily,'H'],[coverage,'G'],[members,'E']])s.getRange(`A1:${last}1`).format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'}};
output.getRange(`B2:E${p.count+1}`).setNumberFormat('#,##0.00');output.getRange('N:N').format.columnWidth=78;
applyCPR(output,p.count+1,'NSE500_'+label,'weekly');
output.getRange('B:E').format.columnWidth=23;
output.getRange('F:J').format.columnWidth=20;
for(let i=0;i<p.calculated.length;i++)if(p.calculated[i][1]===null){
 const row=i+2;output.getRange(`B${row}:K${row}`).values=[['n.a.','n.a.','n.a.','n.a.','n.a.','n.a.','n.a.','n.a.','n.a.','Unavailable']];
}
output.getRange(`K2:K${p.count+1}`).conditionalFormats.add('expression',{formula:'=$K2="Unavailable"',format:{fill:'#FEE2E2',font:{color:'#991B1B',bold:true}}});
daily.getRange(`B2:B${p.daily.length+1}`).setNumberFormat('dd-mmm-yyyy');daily.getRange(`C2:F${p.daily.length+1}`).setNumberFormat('#,##0.00');daily.getRange('G:H').format.columnWidth=24;
coverage.getRange('B:B').format.columnWidth=54;coverage.getRange('G:G').format.columnWidth=22;
coverage.tables.add(`A1:G${p.count+1}`,true,'IndexCoverage');
members.getRange('A:A').format.columnWidth=48;members.getRange('B:B').format.columnWidth=35;members.getRange('E:E').format.columnWidth=22;
members.tables.add(`A1:E${p.count+1}`,true,'CurrentConstituents');
sources.getRange('B:B').format.columnWidth=120;sources.getRange('C:C').format.columnWidth=115;sources.getRange('D:D').format.columnWidth=72;
sources.getRange('B1:B11').format.wrapText=true;sources.getRange('A1:D11').format.rowHeight=42;
sources.getRange('A13:D13').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'}};
const guide=addLegend(w,'weekly');guide.getRange('B9').values=[[p.kind==='monthly'?'Completed calendar-month HLC; levels apply to the following month':'Completed week HLC; levels apply to the following week']];
w.recalculate();
const actual=output.getRange(`A2:L${p.count+1}`).values;
for(let i=0;i<p.count;i++){
 if(actual[i][0]!==p.calculated[i][0])throw Error('Constituent mismatch');
 if(p.calculated[i][1]===null){if(actual[i][10]!=='Unavailable')throw Error('Missing-price entry not exposed');continue;}
 const expected=[...p.calculated[i].slice(1),...expectedCPR(...p.calculated[i].slice(2),'weekly')];
 for(let j=0;j<expected.length;j++){
  const got=actual[i][j+1];if(typeof expected[j]==='number'?Math.abs(got-expected[j])>1e-8:got!==expected[j])throw Error(`NSE500 calculation mismatch ${i},${j}`);
 }
}
verifyTradingView(output,p.count+1);
if(job.render_dir){
 await fs.mkdir(job.render_dir,{recursive:true});
 for(const [s,range] of [[output,'A1:L12'],[daily,'A1:H12'],[coverage,'A1:G12'],[members,'A1:E12'],[sources,'A1:B11'],[guide,'A1:C7']]){
  const im=await w.render({sheetName:s.name,range,scale:1,format:'png'});await fs.writeFile(path.join(job.render_dir,s.name.replaceAll(' ','_')+'.png'),new Uint8Array(await im.arrayBuffer()));
 }
 const missing=p.calculated.findIndex(r=>r[1]===null);
 if(missing>=0){const im=await w.render({sheetName:output.name,range:`A${missing+1}:L${missing+3}`,scale:1,format:'png'});await fs.writeFile(path.join(job.render_dir,'unavailable.png'),new Uint8Array(await im.arrayBuffer()));}
}
await(await SpreadsheetFile.exportXlsx(w)).save(job.output);
console.log(`Verified ${p.count} official index entries, available OHLC/CPR and explicit coverage gaps.`);
