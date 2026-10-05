from pathlib import Path
from datetime import datetime,date
from decimal import Decimal,InvalidOperation
from collections import Counter,OrderedDict
import openpyxl
w=openpyxl.load_workbook('素材_應付明細_台幣.xlsx',read_only=True,data_only=True)
rows=list(w.active.iter_rows(min_row=2,values_only=True));selected=[];invalid=[]
for i,r in enumerate(rows,2):
 try:d=r[7].date() if isinstance(r[7],datetime) else r[7] if isinstance(r[7],date) else datetime.strptime(str(r[7]).strip(),'%Y/%m/%d').date()
 except Exception:invalid.append(i);continue
 if date(2026,10,1)<=d<=date(2026,10,6):selected.append((i,r,d))
counts=Counter(str(r[9]).strip().upper() for r in rows if r[9] and str(r[9]).strip())
groups=OrderedDict();issues=[]
for i,r,d in selected:
 for c in [2,4,9,13,14]:
  if r[c] is None or not str(r[c]).strip():issues.append(f'第 {i} 列必要欄位缺漏')
 if counts[str(r[9]).strip().upper()]>1:issues.append(f'第 {i} 列發票重複：{r[9]}')
 amount=Decimal(str(r[13]));paid=Decimal(str(r[14]))
 if not amount.is_finite() or amount<=0:issues.append(f'第 {i} 列金額無效／零或負數')
 if paid!=0:issues.append(f'第 {i} 列已有付款')
 key=(r[4],r[2]);groups.setdefault(key,[]).append((i,r,d,amount))
total=sum(x[3] for rs in groups.values() for x in rs)
assert total==sum(Decimal(str(r[13])) for _,r,_ in selected)
assert sum(len(rs) for rs in groups.values())==len(selected)
large=[(k,rs) for k,rs in groups.items() if sum(x[3] for x in rs)>1000000]
assert not issues and not invalid and not large
fmt=lambda a:format(a,',.0f')
lines=['# 付款審核卡：2026/10/01–2026/10/06','',
'## 一、本期付款作業風控摘要','',
'| 項目 | 結果 |','|---|---|',
'| 付款到期日區間 | 2026/10/01–2026/10/06，含起訖兩日 |',
'| 所屬付款批次 | 2026 年 10 月第 1 次（1–6 日） |',
'| 實際支付日 | 2026/10/06（使用者指定，尚未執行付款） |',
'| 審核日期 | 2026/10/05 |',
f'| 發票總張數／傳票筆數 | {len(selected)} 張／{len(groups)} 筆 |',
f'| 本期應付總額 | NT$ {fmt(total)} |',
'| 實付確認狀態 | 尚未確認；本期 ERP 本幣已付合計 NT$ 0 |',
'| 巨額筆數／金額 | 0 筆／NT$ 0 |',
'| 異常狀態 | 本次指定檢核未發現異常 |',
'| 審核結論 | 資料檢核通過，待使用者確認，尚未產生 Excel |','',
'## 二、傳票級放行明細清單','',
'| 序號 | 傳票編號 | 廠商代號 | 發票張數 | 本期應付金額（NT$） | 風控評級 | 摘要 |',
'|---:|---|---|---:|---:|---|---|']
for i,((voucher,vendor),rs) in enumerate(groups.items(),1):
 lines.append(f'| {i} | {voucher} | {vendor} | {len(rs)} | {fmt(sum(x[3] for x in rs))} | 檢核通過，待審批 | 應付 {vendor}（{len(rs)}筆發票） |')
lines.extend([f'| **合計** | **{len(groups)} 筆傳票** | — | **{len(selected)}** | **{fmt(total)}** | **待使用者確認** | — |','',
'## 三、巨額单據備忘與異常說明'.replace('单','單'),'',
'本期無單筆歸戶傳票超過 NT$ 1,000,000。各發票原始付款到期日如下：','',
'| 原檔列號 | 傳票編號 | 發票號碼 | 原始到期日 | 本幣應付金額（NT$） |',
'|---:|---|---|---|---:|'])
for i,r,d in selected:lines.append(f'| {i} | {r[4]} | {r[9]} | {d:%Y/%m/%d} | {fmt(Decimal(str(r[13])))} |')
lines.extend(['',
'- 區間核對：2026/10/05 有 6 張、合計 NT$ 1,127,170；2026/10/06 有 2 張、合計 NT$ 517,125。2026/10/01–10/04 無到期款項，未納入區間外資料。',
'- 重複發票：本期 8 張號碼互異；與原檔全部 270 列交叉比對亦未發現重複，未涵蓋其他檔案或歷史付款。',
'- 金額防呆：本期本幣應付均為有效正數，未發現零或負數；本幣已付均為 0。',
'- 完整性：本期傳票號碼、帳款對象、發票號碼與必要金額均完整；全檔付款到期日皆可辨識。',
'- 歸戶核對：8 列／8 張發票歸為 7 筆傳票，明細與歸戶合計均為 NT$ 1,644,295，未漏列或重複計入。',
'- 不驗證原幣與本幣一致性，不計算手續費及銀行餘額。',
'- 合約、驗收、付款里程碑、受款帳號與授權：原檔未提供，待人工核對。','',
'## 四、審批結論與下一步','',
'**資料檢核通過，待使用者確認，尚未產生 Excel。**','',
'請人工核對憑證、歷史付款及受款資料，確認後回覆：**「審核通過，請產生 Excel」**。本期無巨額款項，無需另行確認巨額主管覆核。核准後才套用固定樣板，支付日使用 2026/10/06。','',
'來源：`素材_應付明細_台幣.xlsx`／「應付素材-台幣」／A1:P271；本期 A89:P96。以 H 欄付款到期日作日期區間篩選，E 欄傳票號碼＋C 欄帳款對象歸戶，J 欄發票號碼識別，N 欄本幣應付金額彙總。未修改原始資料及固定樣板。',''])
Path('付款審核卡_20261001_20261006.md').write_text('\n'.join(lines),encoding='utf-8')
print(f'已儲存審核卡片：{len(selected)} 張／{len(groups)} 筆，總額 {fmt(total)}；異常 0、巨額 0。')
