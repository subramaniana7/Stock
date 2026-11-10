import fs from 'node:fs/promises';
import path from 'node:path';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';
import { applyCPR,addLegend,verifyCPR } from './cpr.mjs';
const job=JSON.parse(await fs.readFile(process.argv[2],'utf8'));
const p=job.payload;
const w=Workbook.create();
const dates=Object.keys(p.days).sort();
function nameOf(date){ return new Date(date+'T00:00:00Z').toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'}).replaceAll(' ','-'); }
const sourceRows=[];
for(const date of dates){
 const d=p.days[date]; const s=w.worksheets.add(nameOf(date));
 s.showGridLines=false; s.tabColor='#1F4E78';
 s.getRange('A1:E1').values=[['Stock Symbol','Daily Open','Daily High','Daily Low','Daily Close']];
 s.getRange('A2').write(d.rows);
 s.getRange('M1').values=[[`Trading date: ${date} | INR per share`]];
 s.getRange('M2').values=[[`${d.rows.length} F&O stocks | Unadjusted cash-market prices`]];
 s.getRange(`A1:G${d.rows.length+1}`).format.font={name:'Arial',size:11};
 s.getRange(`A1:E${d.rows.length+1}`).format.rowHeight=21;
 s.getRange('A:A').format.columnWidth=24;
 s.getRange('B:E').format.columnWidth=19;
 s.getRange('L:L').format.columnWidth=3;
 s.getRange('M:M').format.columnWidth=70;
 s.getRange('A1:E1').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'},horizontalAlignment:'center'};
 s.getRange(`B2:E${d.rows.length+1}`).setNumberFormat('#,##0.00');
 s.freezePanes.freezeRows(1);
 applyCPR(s,d.rows.length+1,'Daily_'+date.replaceAll('-',''));
 for(const r of d.sources) sourceRows.push([new Date(r[0]+'T00:00:00Z'),...r.slice(1),'IST '+d.retrieved_at]);
}
const source=w.worksheets.add('Sources'); source.showGridLines=false;
source.getRange('A1:B6').values=[
 ['Report week',`${p.week_start} to ${new Date(new Date(p.week_start+'T00:00:00Z').getTime()+4*86400000).toISOString().slice(0,10)}`],
 ['Prices','INR per share; unadjusted cash-market EQ series / F1 session'],
 ['Universe','Stock underlyings with STF/STO contracts on each individual trading day; index contracts excluded'],
 ['Data checks','Dates, segment, unique symbols, complete cash coverage and OHLC price bounds are validated'],
 ['Sheets','One dated sheet per available trading day; no blank holiday sheets'],
 ['Holiday calendar','https://www.nseindia.com/api/holiday-master?type=trading']];
source.getRange('A8:E8').values=[['Trading Date','Segment','Official NSE source URL','SHA-256','Retrieved at (IST)']];
source.getRange('A9').write(sourceRows);
source.getRange(`A9:A${8+sourceRows.length}`).setNumberFormat('dd-mmm-yyyy');
source.getUsedRange().format.font={name:'Arial',size:11};
source.getUsedRange().format.rowHeight=24;
source.getRange('A:A').format.columnWidth=22; source.getRange('B:B').format.columnWidth=100;
source.getRange('C:C').format.columnWidth=110; source.getRange('D:D').format.columnWidth=72;
source.getRange('E:E').format.columnWidth=50;
source.getRange('A8:E8').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'}};
source.getRange('A1:A6').format.font={bold:true};
addLegend(w,'daily');
w.recalculate();
for(const date of dates){
 const rows=p.days[date].rows;
 const actual=w.worksheets.getItem(nameOf(date)).getRange(`A2:E${rows.length+1}`).values;
 if(JSON.stringify(actual)!==JSON.stringify(rows)) throw Error(`Daily price verification failed for ${date}`);
 verifyCPR(w.worksheets.getItem(nameOf(date)),rows.length+1);
}
const errors=await w.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!',options:{useRegex:true,maxResults:20},maxChars:2000});
if(/"cell"/.test(errors.ndjson)) throw Error(errors.ndjson);
if(job.render_dir){
 await fs.mkdir(job.render_dir,{recursive:true});
 for(const date of dates){
  const preview=await w.render({sheetName:nameOf(date),range:'A1:K12',scale:1.2,format:'png'});
  await fs.writeFile(path.join(job.render_dir,date+'.png'),new Uint8Array(await preview.arrayBuffer()));
 }
 const preview=await w.render({sheetName:'Sources',range:'A1:B6',scale:1.1,format:'png'});
 await fs.writeFile(path.join(job.render_dir,'sources.png'),new Uint8Array(await preview.arrayBuffer()));
 const guide=await w.render({sheetName:'CPR Guide',range:'A1:C5',scale:1,format:'png'});
 await fs.writeFile(path.join(job.render_dir,'cpr-guide.png'),new Uint8Array(await guide.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(w)).save(job.output);
console.log(`Verified ${dates.length} daily sheets and exported ${job.output}`);
