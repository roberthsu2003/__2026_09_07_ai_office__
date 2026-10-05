import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';
const w=await SpreadsheetFile.importXlsx(await FileBlob.load('outputs/01a10bc7-8ae6-7fc0-82fa-7915b40e0642/出納付款放行總表_已完成_Excel相容版.xlsx'));
w.recalculate();
const s=w.worksheets.getItemAt(0);
if(s.getRange('H9').values[0][0]!==4040257)throw Error('計算錯誤');
const p=await w.render({sheetName:s.name,range:'A1:J12',scale:1});
await fs.writeFile('.work/compatible.png',new Uint8Array(await p.arrayBuffer()));
console.log('另一讀取引擎驗證資料及公式成功');
