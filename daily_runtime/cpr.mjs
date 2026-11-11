export const CPR_HEADERS=['Pivot','TC','BC','TC-BC','CPR Width %','CPR Classification'];
export function classificationFormula(period='daily'){
 // Tiny tolerance prevents binary floating-point ties (e.g. 0.5005) rounding down.
 const rounded='ROUND(J2+1E-12,3)';
 const wider=period==='weekly'?`IF(${rounded}<=0.500,"Medium 1",IF(${rounded}<=0.800,"Medium 2","Above Medium"))`:'"Above Medium"';
 return `=IF(ISNUMBER(J2),IF(${rounded}<=0.050,"Very Narrow",IF(${rounded}<=0.100,"Narrow",IF(${rounded}<=0.200,"Medium",${wider}))),"")`;
}
export const CLASS_FORMULA=classificationFormula();
export const CLASS_COLORS={'Very Narrow':'#DCFCE7','Narrow':'#DBEAFE','Medium':'#FEF3C7','Medium 1':'#FFEDD5','Medium 2':'#EDE9FE','Above Medium':'#F3F4F6'};
export function classifyWidth(value,period='daily'){
 if(typeof value!=='number')return '';
 const rounded=Math.round((value+1e-12)*1000)/1000;
 return rounded<=.050?'Very Narrow':rounded<=.100?'Narrow':rounded<=.200?'Medium':period==='weekly'?(rounded<=.500?'Medium 1':rounded<=.800?'Medium 2':'Above Medium'):'Above Medium';
}
export function updateClassification(sheet,lastRow,period='daily'){
 sheet.getRange('K2').formulas=[[classificationFormula(period)]];
 sheet.getRange(`K2:K${lastRow}`).fillDown();
 sheet.getRange(`J2:J${lastRow}`).setNumberFormat('0.000');
 const names=period==='weekly'?Object.keys(CLASS_COLORS):['Very Narrow','Narrow','Medium','Above Medium'];
 sheet.getRange(`K2:K${lastRow}`).conditionalFormats.deleteAll();
 for(const label of names)sheet.getRange(`K2:K${lastRow}`).conditionalFormats.add('expression',{
  formula:`=$K2="${label}"`,format:{fill:CLASS_COLORS[label],font:{color:'#1F2937',bold:true}}});
}
export function expectedCPR(high,low,close,period='daily'){
 const pivot=(high+low+close)/3, midpoint=(high+low)/2, other=2*pivot-midpoint;
 const tc=Math.max(midpoint,other),bc=Math.min(midpoint,other),width=tc-bc;
 const percent=pivot===0?'':width/pivot*100;
 const label=classifyWidth(percent,period);
 return [pivot,tc,bc,width,percent,label];
}
export function applyCPR(sheet,lastRow,tableName,period='daily'){
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
 updateClassification(sheet,lastRow,period);
 addTradingViewColumn(sheet,lastRow,tableName);
}
export function addTradingViewColumn(sheet,lastRow,tableName){
 sheet.getRange('L1').values=[['TradingView Symbol']];
 sheet.getRange('L2').formulas=[['=A2&","']];
 sheet.getRange(`L2:L${lastRow}`).fillDown();
 sheet.getRange(`L1:L${lastRow}`).format.font={name:'Arial',size:11};
 sheet.getRange('L:L').format.columnWidth=24;
 sheet.getRange('L1').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'},horizontalAlignment:'center'};
 sheet.getRange(`L2:L${lastRow}`).setNumberFormat('General');
 const tables=sheet.tables.items;
 let name=tableName,style;
 for(const table of [...tables]){name=table.name;style=table.style;table.delete();}
 const table=sheet.tables.add(`A1:L${lastRow}`,true,name);
 if(style) table.style=style;
 table.showFilterButton=true;
}
export function addLegend(w,period){
 return writeLegend(w.worksheets.add('CPR Guide'),period);
}
export function writeLegend(s,period){
 s.getUsedRange()?.clear({applyTo:'all'});s.showGridLines=false;
 const bands=[
  ['Classification','CPR Width %','Possible behaviour'],
  ['Very Narrow','0.000% through 0.050%','Possible breakout or strong directional expansion'],
  ['Narrow','0.051% through 0.100%','Potential expansion from a tight range'],
  ['Medium','0.101% through 0.200%','Moderate or transitional movement']];
 if(period==='weekly')bands.push(['Medium 1','0.201% through 0.500%','Wider CPR band; interpret alongside trend and volatility'],['Medium 2','0.501% through 0.800%','Wider CPR band; interpret alongside trend and volatility']);
 bands.push(['Above Medium',period==='weekly'?'0.801% and above':'0.201% and above','Wider CPR range; interpret alongside price trend and volatility']);
 s.getRange(`A1:C${bands.length}`).values=bands;
 s.getRange('A9:B18').values=[
  ['Reference period',period==='weekly'?'Completed week HLC; levels apply to the following week':'Completed session HLC; levels apply to the next trading session'],
  ['Pivot','(High + Low + Close) / 3'],
  ['Central boundaries','Midpoint = (High + Low) / 2; other boundary = 2 × Pivot − midpoint'],
  ['TC / BC','TC is the higher boundary; BC is the lower boundary'],
  ['Width','TC − BC; nonnegative'],
  ['CPR Width %','(TC − BC) / Pivot × 100; 0.05 in column J means 0.05%'],
  ['Classification','Uses ROUND(CPR Width %, 3); displayed width and classification use the same three-decimal value'],
  ['Precision','CPR calculations retain full precision; Pivot displays whole rupees; TC/BC/width two decimals; width % three decimals'],
  ['Interpretation','These thresholds are descriptive categories; market behaviour also depends on other factors'],
  ['Formula source','https://support.zerodha.com/category/trading-and-markets/charts-and-orders/charts/articles/how-are-pivot-points-derived-while-using-the-cpr']];
 s.getRange('A1:C18').format.font={name:'Arial',size:11};
 s.getRange('A1:C18').format.rowHeight=27;
 s.getRange('A:A').format.columnWidth=24;s.getRange('B:B').format.columnWidth=105;s.getRange('C:C').format.columnWidth=80;
 s.getRange('A1:C1').format={fill:'#1F4E78',font:{bold:true,color:'#FFFFFF'}};
 for(let i=1;i<bands.length;i++)s.getRange(`A${i+1}`).format.fill=CLASS_COLORS[bands[i][0]];
 return s;
}
export function verifyCPR(sheet,lastRow,period='daily'){
 const input=sheet.getRange(`C2:E${lastRow}`).values;
 const actual=sheet.getRange(`F2:K${lastRow}`).values;
 for(let i=0;i<input.length;i++){
  const expected=expectedCPR(...input[i],period);
  for(let j=0;j<6;j++){
   const equal=typeof expected[j]==='number'?typeof actual[i][j]==='number'&&Math.abs(actual[i][j]-expected[j])<1e-8:actual[i][j]===expected[j];
   if(!equal) throw Error(`CPR mismatch ${sheet.name} row ${i+2}, metric ${j}: ${actual[i][j]} versus ${expected[j]}`);
  }
 }
 return input.length;
}
export function verifyTradingView(sheet,lastRow){
 const symbols=sheet.getRange(`A2:A${lastRow}`).values;
 const actual=sheet.getRange(`L2:L${lastRow}`).values;
 for(let i=0;i<symbols.length;i++){
  if(actual[i][0]!==symbols[i][0]+',') throw Error(`TradingView symbol mismatch ${sheet.name}!L${i+2}`);
 }
 return symbols.length;
}
