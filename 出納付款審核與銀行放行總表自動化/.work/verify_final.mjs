import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const w=await SpreadsheetFile.importXlsx(await FileBlob.load('outputs/01a10bc7-8ae6-7fc0-82fa-7915b40e0642/出納付款放行總表_已完成.xlsx'));
w.recalculate();
const s=w.worksheets.getItemAt(0);
if(s.getRange('F9').values[0][0]!==4040257||s.getRange('H9').values[0][0]!==4040257)throw Error('總額錯誤');
const p=await w.render({sheetName:s.name,range:'A1:J13',scale:1.3});
await fs.writeFile('.work/completed-verified.png',new Uint8Array(await p.arrayBuffer()));
console.log('最終檔案重新載入及總額驗證通過。');
