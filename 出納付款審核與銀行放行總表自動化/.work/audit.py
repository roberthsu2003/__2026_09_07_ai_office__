import zipfile,xml.etree.ElementTree as E,json
from decimal import Decimal
from collections import defaultdict,Counter
from pathlib import Path
n={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile('素材_應付明細_台幣.xlsx') as z:
 ss=[''.join(t.itertext()) for t in E.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si',n)]
 rows=[]
 for r in E.fromstring(z.read('xl/worksheets/sheet1.xml')).findall('m:sheetData/m:row',n):
  d={'row':int(r.attrib['r'])}
  for c in r.findall('m:c',n):
   v=c.find('m:v',n)
   if v is not None: d[''.join(x for x in c.attrib['r'] if x.isalpha())]=ss[int(v.text)].strip() if c.attrib.get('t')=='s' else v.text
  if d['row']>1 and d.get('H'): rows.append(d)
selected=[r for r in rows if r['H']=='2026/09/30']
groups=defaultdict(list)
for r in selected: groups[(r.get('E',''),r.get('C',''))].append(r)
counts=Counter(r.get('J','').upper() for r in rows if r.get('J'))
dups={k:v for k,v in counts.items() if v>1 and any(r.get('J','').upper()==k for r in selected)}
bad=[r for r in selected if Decimal(r['M'])!=Decimal(r['N'])]
nonpositive=[r for r in selected if any(Decimal(r[c])<=0 for c in ['K','M','N','P'])]
print(json.dumps({'source_rows':len(rows),'selected_rows':len(selected),'groups':len(groups),'duplicates':dups,'mismatch':bad,'nonpositive':nonpositive,'missing_invoice':[r['row'] for r in selected if not r.get('J')],'currencies':dict(Counter(r['F'] for r in selected)),'paid_sum':str(sum(Decimal(r['O']) for r in selected))},ensure_ascii=False))
for (v,c),rs in groups.items(): print(json.dumps({'voucher':v,'vendor':c,'n':len(rs),'amount':str(sum(Decimal(r['N']) for r in rs)),'rows':[r['row'] for r in rs],'invoices':[r['J'] for r in rs],'dates':sorted(set(r['A'] for r in rs))},ensure_ascii=False))
money=lambda x:format(x,',.2f') if x%1 else format(x,',.0f')
total=sum(Decimal(r['N']) for r in selected)
blocked=sum(Decimal(r['N']) for r in selected if (r['E'],r['C']) in {(x['E'],x['C']) for x in bad})
large=[(key,rs) for key,rs in groups.items() if sum(Decimal(r['N']) for r in rs)>1000000]
assert len(selected)==sum(len(rs) for rs in groups.values())
assert total==sum(sum(Decimal(r['N']) for r in rs) for rs in groups.values())
lines=['# 出納付款風控審核與傳票歸戶卡','',
'## 1. 本期付款作業風控摘要','',
'| 項目 | 審核結果 |','|---|---|',
'| 付款到期日／統計基準日 | 2026/09/30 |',
'| 審核日期 | 2026/10/05 |',
f'| 發票總張數 | {len(selected)} 張，發票號碼均有值且互異 |',
f'| 傳票筆數 | {len(groups)} 筆（傳票號碼＋帳款對象） |',
f'| 本期應付／預計付款總額 | NT$ {money(total)}，未含銀行手續費 |',
'| 實付總額 | 尚無銀行付款證據；ERP 本幣已付金額合計為 NT$ 0 |',
f'| 巨額傳票 | {len(large)} 筆，NT$ {money(sum(sum(Decimal(r["N"]) for r in rs) for _,rs in large))}（嚴格超過 NT$ 1,000,000） |',
f'| 異常狀態 | 沖平不符 {len(bad)} 張／3 筆傳票，暫緩金額 NT$ {money(blocked)} |',
f'| 通過本階段資料檢核 | 2 筆／2 張，NT$ {money(total-blocked)}；仍待人工審批 |',
'',
'檢核口徑：來源為「素材_應付明細_台幣.xlsx」工作表「應付素材-台幣」，270 列資料中以 H 欄付款到期日精確篩選，命中第 79–88 列。E 欄「傳票號碼」對應傳票編號，C 欄「帳款對象」對應廠商代號，J 欄「發票號碼」作為發票識別（I 欄發票編號為空白）。以 N 欄本幣應付金額彙總；本期 O 欄已付均為零，故 N 欄與 P 欄本幣未付逐列相等。',
'',
'## 2. 傳票級放行明細清單','',
'| 序號 | 傳票編號 | 廠商代號 | 發票張數 | 本期應付金額（NT$） | 風控評級 | 摘要 |',
'|---:|---|---|---:|---:|---|---|']
for i,((v,c),rs) in enumerate(groups.items(),1):
 amount=sum(Decimal(r['N']) for r in rs)
 risk='沖平不符，暫緩放行' if any(r in bad for r in rs) else '資料檢核通過，待人工審批'
 if amount>1000000:risk='🚨 巨額支付特別覆核；'+risk
 lines.append(f'| {i} | {v} | {c} | {len(rs)} | {money(amount)} | {risk} | 應付 {c}（{len(rs)}筆發票） |')
lines += [f'| **合計** | **5 筆傳票** | — | **10** | **{money(total)}** | **3 筆暫緩／2 筆待審批** | — |','',
'### 防呆與沖平結果','',
'- 重複發票號碼：本期未發現；將本期號碼與原檔全部 270 列交叉比對，亦未發現重複。此結果未涵蓋其他檔案或歷史付款紀錄。',
'- 零或負數：本期 K／M／N／P 欄應付與未付金額皆大於零；L／O 欄已付金額為零屬未付款狀態。',
'- 歸戶鍵、發票號碼與必要金額：本期皆完整；幣別皆標示 TWD。',
'- 沖平：逐列以十進位精確比較 M 欄原幣未付與 N 欄本幣應付，2 張相等、8 張不相等；不得以彙總相抵取代逐筆檢核。','',
'| 原檔列號 | 傳票編號 | 發票號碼 | 原幣未付 | 本幣應付（NT$） | 差額（本幣－原幣） |',
'|---:|---|---|---:|---:|---:|']
for r in bad:lines.append(f'| {r["row"]} | {r["E"]} | {r["J"]} | {money(Decimal(r["M"]))} | {money(Decimal(r["N"]))} | {money(Decimal(r["N"])-Decimal(r["M"]))} |')
lines += ['',
'差額僅供定位欄位不一致，不代表應調帳金額。部分數值呈現換算關係，但原檔幣別標示 TWD，無匯率與換算依據，不能認定為合理匯兌；應由會計確認幣別、ERP 匯出欄位及原始憑證。','',
'## 3. 巨額單據備忘','',
'**🚨 巨額支付特別覆核：F1-GL01-260722013／TB0001，2 張發票，NT$ 3,685,681。** 立帳日 2026/07/22，付款到期日 2026/09/30，原檔第 85–86 列。','',
'| 發票號碼 | 原幣未付 | 本幣應付（NT$） | 狀態 |','|---|---:|---:|---|',
'| CA36795506 | 99,009.62 | 3,153,459 | 沖平不符 |',
'| CA36795519 | 16,710.23 | 532,222 | 沖平不符 |',
'| **合計** | **115,719.85** | **3,685,681** | **暫緩放行** |','',
'合約期程提醒：原檔未提供合約、付款里程碑、驗收或保留款條款。請核對兩張發票所屬合約與採購單，確認分期比例、驗收條件、保留款／預付款扣抵及到期日，避免重複支付或提前付款。截至審核日已超過到期日 5 天；須確認是否已付款、展延或存在逾期費用，再依公司授權層級完成巨額覆核。','',
'## 4. 審批結論與人機協作下一步','',
'**本期不具備整批放行條件。** 3 筆傳票合計 NT$ 4,036,870 因沖平不符暫緩，其中巨額傳票另須特別覆核。其餘 2 筆合計 NT$ 3,387 僅通過本階段資料檢核，尚未取得實際付款授權。','',
'1. **會計／ERP 負責人：**釐清上述 8 張發票的幣別與金額差異，修正 ERP 或提供可核驗的換算依據，回傳更新明細後重新執行檢核。',
'2. **出納／複核人：**核對全期歷史付款、銀行回執、受款戶名帳號、發票及驗收證明。對 2 筆通過資料檢核的傳票完成審批；巨額傳票須先解決沖平差異，再完成合約與授權覆核。',
'3. **人機協作：**人工確認合格清單、付款銀行、付款日、期初餘額與手續費後，再以已建立的佔位字符樣版產生第二階段放行總表。不得將本期應付合計直接標為銀行實付，也不得將暫緩傳票混入付款批次。','',
'來源：素材_應付明細_台幣.xlsx／應付素材-台幣／A1:P271；本期明細 A79:P88。金額歸戶與逐列沖平均使用原檔儲存值，未改動來源檔。','']
out=Path('outputs/01a10bc7-8ae6-7fc0-82fa-7915b40e0642/20260930_出納付款風控審核卡.md')
out.write_text('\n'.join(lines),encoding='utf-8')
print('SAVED',out,'TOTAL',total,'BLOCKED',blocked,'PASS',total-blocked)
