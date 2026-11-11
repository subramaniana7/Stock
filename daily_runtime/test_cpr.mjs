import {Workbook} from '@oai/artifact-tool';
import {applyCPR,verifyCPR,classificationFormula} from './cpr.mjs';
const w=Workbook.create(),s=w.worksheets.add('CPR Tests');
s.getRange('A1:E1').values=[['Symbol','Open','High','Low','Close']];
s.getRange('A2').write([['TOP',100,110,90,110],['REVERSE',100,110,90,90],['FLAT',100,100,100,100],['ZERO',0,0,0,0]]);
applyCPR(s,5,'TestPrices');w.recalculate();verifyCPR(s,5);
if(s.getRange('G3').values[0][0]<s.getRange('H3').values[0][0])throw Error('TC below BC');
const widths=[0,.050,.0504,.0505,.051,.100,.1004,.1005,.101,.200,.2004,.2005,.201,.500,.5004,.5005,.501,.800,.8004,.8005,.801,null];
const daily=['Very Narrow','Very Narrow','Very Narrow','Narrow','Narrow','Narrow','Narrow','Medium','Medium','Medium','Medium','Above Medium','Above Medium','Above Medium','Above Medium','Above Medium','Above Medium','Above Medium','Above Medium','Above Medium','Above Medium',''];
const weekly=['Very Narrow','Very Narrow','Very Narrow','Narrow','Narrow','Narrow','Narrow','Medium','Medium','Medium','Medium','Medium 1','Medium 1','Medium 1','Medium 1','Medium 2','Medium 2','Medium 2','Medium 2','Above Medium','Above Medium',''];
for(const period of ['daily','weekly']){
 const b=w.worksheets.add(period+' Boundaries');
 b.getRange('J2').write(widths.map(x=>[x]));
 b.getRange('K2').formulas=[[classificationFormula(period)]];b.getRange(`K2:K${widths.length+1}`).fillDown();w.recalculate();
 const actual=b.getRange(`K2:K${widths.length+1}`).values.map(r=>r[0]);
 if(JSON.stringify(actual)!==JSON.stringify(period==='daily'?daily:weekly))throw Error(period+': '+JSON.stringify(actual));
}
console.log('PASS: 44 daily/weekly classification cases including inclusive endpoints, 3-decimal rounding transitions, 0.201/0.801, blanks, normalized CPR and zero pivot.');
