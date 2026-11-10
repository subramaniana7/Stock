export const CPR_HEADERS=['Pivot','TC','BC','TC-BC','CPR Width %','CPR Classification'];
export const CLASS_FORMULA='=IF(ISNUMBER(J2),IF(J2<0.05,"Very Narrow",IF(J2<0.10,"Narrow",IF(J2<=0.20,"Medium","Above Medium"))),"")';
export const CLASS_COLORS={'Very Narrow':'#DCFCE7','Narrow':'#DBEAFE','Medium':'#FEF3C7','Above Medium':'#F3F4F6'};
export function expectedCPR(high,low,close){
 const pivot=(high+low+close)/3, midpoint=(high+low)/2, other=2*pivot-midpoint;
 const tc=Math.max(midpoint,other),bc=Math.min(midpoint,other),width=tc-bc;
 const percent=pivot===0?'':width/pivot*100;
 const label=percent===''?'':percent<0.05?'Very Narrow':percent<0.10?'Narrow':percent<=0.20?'Medium':'Above Medium';
 return [pivot,tc,bc,width,percent,label];
}
export function applyCPR(sheet,lastRow,tableName){
 sheet.getRange('F1:K1').values=[CPR_HEADERS];
 sheet.getRange('F2:K2').formulas=[[
  '=(C2+D2+E2)/3',
  '=MAX((C2+D2)/2,2*F2-(C2+D2)/2)',
  '=MIN((C2+D2)/2,2*F2-(C2+D2)/2)',
  '=G2-H2','=IF(F2=0,"",I2/F2*100)',CLASS_FORMULA]];
 sheet.getRange(`F2:K${lastRow}`).fillDown();
 sheet.getRange(`F1:K${lastRow}`).format.font={name:'Arial',size:11};
 sheet.getRange('F:H').format.columnWidth=16;
 sheet.getRange('I:I').format.columnWidth=15;
 sheet.getRange('J:J').format.columnWidth=18;
 sheet.getRange('K:K').format.columnWidth=24;
 sheet.getRange('F1:K1').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'},horizontalAlignment:'center'};
 sheet.getRange(`F2:F${lastRow}`).setNumberFormat('#,##0');
 sheet.getRange(`G2:J${lastRow}`).setNumberFormat('#,##0.00');
 for(const [label,color] of Object.entries(CLASS_COLORS)){
  sheet.getRange(`K2:K${lastRow}`).conditionalFormats.add('expression',{
   formula:`=$K2="${label}"`,format:{fill:color,font:{color:'#1F2937',bold:true}}});
 }
 const tables=sheet.tables.items;
 let name=tableName,style;
 for(const table of [...tables]){name=table.name;style=table.style;table.delete();}
 const table=sheet.tables.add(`A1:K${lastRow}`,true,name);
 if(style) table.style=style;
 table.showFilterButton=true;
}
export function addLegend(w,period){
 const s=w.worksheets.add('CPR Guide'); s.showGridLines=false;
 s.getRange('A1:C5').values=[
  ['Classification','CPR Width %','Possible behaviour'],
  ['Very Narrow','Below 0.05%','Possible breakout or strong directional expansion'],
  ['Narrow','0.05% to below 0.10%','Potential expansion from a tight range'],
  ['Medium','0.10% through 0.20%','Moderate or transitional movement'],
  ['Above Medium','Above 0.20%','Wider CPR range; interpret alongside price trend and volatility']];
 s.getRange('A7:B16').values=[
  ['Reference period',period==='weekly'?'Completed week HLC; levels apply to the following week':'Completed session HLC; levels apply to the next trading session'],
  ['Pivot','(High + Low + Close) / 3'],
  ['Central boundaries','Midpoint = (High + Low) / 2; other boundary = 2 × Pivot − midpoint'],
  ['TC / BC','TC is the higher boundary; BC is the lower boundary'],
  ['Width','TC − BC; nonnegative'],
  ['CPR Width %','(TC − BC) / Pivot × 100; 0.05 in column J means 0.05%'],
  ['Classification','Uses the unrounded value; displayed two-decimal widths can conceal boundary differences'],
  ['Precision','Full precision in formulas; Pivot displays whole rupees, TC/BC/width display two decimals'],
  ['Interpretation','These thresholds are descriptive categories; market behaviour also depends on other factors'],
  ['Formula source','https://support.zerodha.com/category/trading-and-markets/charts-and-orders/charts/articles/how-are-pivot-points-derived-while-using-the-cpr']];
 s.getRange('A1:C16').format.font={name:'Arial',size:11};
 s.getRange('A1:C16').format.rowHeight=27;
 s.getRange('A:A').format.columnWidth=24;s.getRange('B:B').format.columnWidth=105;s.getRange('C:C').format.columnWidth=80;
 s.getRange('A1:C1').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'}};
 for(const [i,label] of ['Very Narrow','Narrow','Medium','Above Medium'].entries()) s.getRange(`A${i+2}`).format.fill=CLASS_COLORS[label];
 return s;
}
export function verifyCPR(sheet,lastRow){
 const input=sheet.getRange(`C2:E${lastRow}`).values;
 const actual=sheet.getRange(`F2:K${lastRow}`).values;
 for(let i=0;i<input.length;i++){
  const expected=expectedCPR(...input[i]);
  for(let j=0;j<6;j++){
   const equal=typeof expected[j]==='number'?typeof actual[i][j]==='number'&&Math.abs(actual[i][j]-expected[j])<1e-8:actual[i][j]===expected[j];
   if(!equal) throw Error(`CPR mismatch ${sheet.name} row ${i+2}, metric ${j}: ${actual[i][j]} versus ${expected[j]}`);
  }
 }
 return input.length;
}
