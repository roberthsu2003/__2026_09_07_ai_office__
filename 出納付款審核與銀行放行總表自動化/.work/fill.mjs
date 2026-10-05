import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load('放行總表_佔位字符樣版.xlsx'));
const s=wb.worksheets.getItemAt(0);
const p=await wb.render({sheetName:s.name,range:'A1:J12',scale:1});
await fs.writeFile('.work/fill-source.png',new Uint8Array(await p.arrayBuffer()));
const erp=await SpreadsheetFile.importXlsx(await FileBlob.load('素材_應付明細_台幣.xlsx'));
const rows=erp.worksheets.getItemAt(0).getRange('A2:P271').values;
const groups=new Map();
for(const row of rows){
 if(String(row[7]).trim()!=='2026/09/30')continue;
 const key=String(row[4]).trim()+'|'+String(row[2]).trim();
 if(!groups.has(key))groups.set(key,{voucher:String(row[4]).trim(),vendor:String(row[2]).trim(),count:0,amount:0});
 const g=groups.get(key);g.count++;g.amount+=Number(row[13]);
}
const data=[...groups.values()];
if(data.length!==5||data.reduce((a,g)=>a+g.amount,0)!==4040257)throw Error('歸戶驗證失敗');
const end=4+data.length,total=end+1,note=total+1,sign=total+2;
s.unmergeCells('A9:D9');s.unmergeCells('A11:J11');
const delta=data.length-4;
for(let r=12;r>=9;r--)s.getRange(`A${r+delta}:O${r+delta}`).copyFrom(s.getRange(`A${r}:O${r}`),'all');
s.getRange(`A${total}:D${total}`).merge();s.getRange(`A${sign}:J${sign}`).merge();
for(let r=5;r<=end;r++){
 s.getRange(`A${r}:J${r}`).copyFrom(s.getRange('A4:J4'),'all');
 s.getRange(`A${r}:J${r}`).clear({applyTo:'contents'});
 s.getRange(`A${r}:J${r}`).formulas=[Array(10).fill('')];
}
s.getRange('A2').values=[['臺企銀 臺幣單筆付款 放行總表']];
s.getRange('J2').values=[['FIN115-175']];
s.getRange('D4').values=[['2026/09/30-存款餘額（待填）']];
s.getRange('I4').clear({applyTo:'contents'});
s.getRange('I4').format.fill='#FFF2CC';
for(let i=0;i<data.length;i++){
 const r=i+5,g=data[i];
 s.getRange(`A${r}:J${r}`).values=[[i+1,g.voucher,new Date('2026-09-30T00:00:00Z'),`應付 ${g.vendor}（${g.count}筆發票）`,0,g.amount,null,null,null,null]];
 s.getRange(`C${r}`).setNumberFormat('yyyy/mm/dd');
 s.getRange(`H${r}`).formulas=[[`=IF(ISNUMBER(G${r}),SUM(F${r}:G${r}),"")`]];
 s.getRange(`I${r}`).formulas=[[`=IF(AND(ISNUMBER(I${r-1}),ISNUMBER(H${r})),I${r-1}+E${r}-H${r},"")`]];
 s.getRange(`G${r}`).format.fill='#FFF2CC';
}
s.getRange(`A${total}`).values=[['合    計：']];
for(const c of ['E','F'])s.getRange(`${c}${total}`).formulas=[[`=SUM(${c}5:${c}${end})`]];
for(const c of ['G','H'])s.getRange(`${c}${total}`).formulas=[[`=IF(COUNT(${c}5:${c}${end})=${data.length},SUM(${c}5:${c}${end}),"")`]];
s.getRange(`I${total}`).formulas=[[`=IF(ISNUMBER(I${end}),I${end},"")`]];
s.getRange(`E4:I${total}`).setNumberFormat('#,##0');
s.getRange(`A${note}:J${note}`).clear({applyTo:'contents'});
s.getRange(`A${note}:J${note}`).merge();
s.getRange(`A${note}`).values=[['本期手續費由公司負擔；單筆大額款項請主管覆核放行。']];
s.getRange(`A${note}:J${note}`).format.font.size=12;
s.getRange(`A${note}:J${note}`).format.horizontalAlignment='left';
s.getRange(`A${sign}`).values=[['核准：                         覆核：                         審核：                         經辦：']];
s.getRange('B1').values=[['手續費與期初餘額待填（黃色欄位）；無手續費請填 0。']];
s.getRange('B1:J1').format.font.size=10;
// Verify the editable fee and opening balance drive all downstream formulas.
s.getRange('I4').values=[[5000000]];
s.getRange(`G5:G${end}`).values=Array.from({length:data.length},()=>[30]);
wb.recalculate();
console.log('TEST',JSON.stringify(s.getRange(`E4:I${total}`).values),JSON.stringify(s.getRange(`H${total}:I${total}`).formulas));
if(s.getRange(`H${total}`).values[0][0]!==4040407||s.getRange(`I${total}`).values[0][0]!==959593)throw Error('手續費／結餘驗證失敗');
s.getRange('I4').clear({applyTo:'contents'});s.getRange(`G5:G${end}`).clear({applyTo:'contents'});
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:`'${s.name}'!A4:J${total}`,include:'values,formulas',tableMaxRows:8,tableMaxCols:10,maxChars:2500})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|\\{\\{',options:{useRegex:true,maxResults:20},maxChars:1000})).ndjson);
const out='outputs/01a10bc7-8ae6-7fc0-82fa-7915b40e0642/出納付款放行總表_已完成.xlsx';
await (await SpreadsheetFile.exportXlsx(wb)).save(out);
const preview=await wb.render({sheetName:s.name,range:`A1:J${sign+1}`,scale:1.3});
await fs.writeFile('.work/completed.png',new Uint8Array(await preview.arrayBuffer()));
