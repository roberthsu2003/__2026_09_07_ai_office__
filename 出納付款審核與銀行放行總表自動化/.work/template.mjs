import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load('素材_放行總表.xlsx'));
console.log((await wb.inspect({kind:'sheet',include:'id,name',maxChars:2000})).ndjson);
const s=wb.worksheets.getItemAt(0);
console.log(JSON.stringify(s.getRange('A1:J12').values));
const p=await wb.render({sheetName:s.name,range:'A1:J12',scale:1.3});
await fs.writeFile('.work/source.png',new Uint8Array(await p.arrayBuffer()));
s.getRange('A2').values=[['{{銀行名稱}} {{幣別}} {{付款方式}} 放行總表']];
s.getRange('J2').values=[['{{表單編號}}']];
s.getRange('J2').format.font.size=10;
s.getRange('D4').values=[['{{餘額日期}}-存款餘額({{帳號末碼}})']];
s.getRange('I4').values=[['{{期初餘額}}']];
const fields=['序號','傳票編號','支付日','摘要','本期收入','本期未付','手續費','本期應付含手續費','結餘','銀行交易序號'];
for(let r=5;r<=8;r++) {
  s.getRange(`A${r}:J${r}`).values=[fields.map(f=>`{{${f}_${r-4}}}`)];
}
s.getRange('A5:J8').format.font.size=10;
s.getRange('A5:J8').format.wrapText=true;
s.getRange('A5:J8').format.verticalAlignment='center';
s.getRange('A11').values=[['核准：{{核准人}}          複核：{{複核人}}          審核：{{審核人}}          經辦：{{經辦人}}']];
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:`'${s.name}'!A2:J9`,include:'values,formulas',tableMaxRows:8,tableMaxCols:10,maxChars:2500})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!',options:{useRegex:true,maxResults:20},maxChars:1500})).ndjson);
const finalPreview=await wb.render({sheetName:s.name,range:'A1:J12',scale:1.3});
await fs.writeFile('.work/final.png',new Uint8Array(await finalPreview.arrayBuffer()));
await (await SpreadsheetFile.exportXlsx(wb)).save('outputs/01a10bc7-8ae6-7fc0-82fa-7915b40e0642/放行總表_佔位字符樣版.xlsx');
