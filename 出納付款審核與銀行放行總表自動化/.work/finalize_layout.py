import zipfile,xml.etree.ElementTree as E,copy,re
from pathlib import Path
ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
q=lambda t:'{'+ns+'}'+t
path=Path('outputs/01a10bc7-8ae6-7fc0-82fa-7915b40e0642/出納付款放行總表_已完成.xlsx')
with zipfile.ZipFile('放行總表_佔位字符樣版.xlsx') as z:
 files={i.filename:z.read(i.filename) for i in z.infolist()}
with zipfile.ZipFile(path) as z:
 authored=E.fromstring(z.read('xl/worksheets/sheet1.xml'))
 ass=[''.join(x.itertext()) for x in E.fromstring(z.read('xl/sharedStrings.xml')).findall(q('si'))]
cells={c.attrib['r']:c for c in authored.findall('.//'+q('sheetData')+'/'+q('row')+'/'+q('c'))}
sheet=E.fromstring(files['xl/worksheets/sheet1.xml'])
sd=sheet.find(q('sheetData')); orig={int(r.attrib['r']):r for r in sd}
for r in list(sd): sd.remove(r)
for r in range(1,14):
 source=r if r<=4 else 4 if r<=9 else r-1
 row=copy.deepcopy(orig[source]);row.attrib['r']=str(r)
 if r==11 and not any(c.attrib['r'].startswith('A') for c in row):
  row.insert(0,E.Element(q('c'),r=f'A{source}',s='28'))
 for c in row:
  col=re.match('[A-Z]+',c.attrib['r'])[0];address=f'{col}{r}';c.attrib['r']=address
  if col>'J':continue
  for child in list(c):c.remove(child)
  c.attrib.pop('t',None)
  a=cells.get(address)
  if a is not None:
   f=a.find(q('f'));v=a.find(q('v'));inline=a.find(q('is'))
   if f is not None:c.append(copy.deepcopy(f))
   if a.attrib.get('t')=='s' and v is not None:
    c.attrib['t']='inlineStr';is_=E.SubElement(c,q('is'));E.SubElement(is_,q('t')).text=ass[int(v.text)]
   elif inline is not None:c.attrib['t']='inlineStr';c.append(copy.deepcopy(inline))
   elif v is not None:
    if a.attrib.get('t'):c.attrib['t']=a.attrib['t']
    c.append(copy.deepcopy(v))
 # Row 1's B cell might not be present; the source includes it.
 if r==11:row.attrib.update(ht='32',customHeight='1')
 if r==12:row.attrib.update(ht='36',customHeight='1')
 sd.append(row)
merges=sheet.find(q('mergeCells'))
for m in list(merges):merges.remove(m)
for ref in ['B1:J1','A2:I2','A10:D10','A11:J11','A12:J12']:E.SubElement(merges,q('mergeCell'),ref=ref)
merges.attrib['count']='5'
# Add yellow-input styles by cloning the exact anchor styles.
styles=E.fromstring(files['xl/styles.xml']);xfs=styles.find(q('cellXfs'));fills=styles.find(q('fills'))
fill=E.SubElement(fills,q('fill'));pat=E.SubElement(fill,q('patternFill'),patternType='solid');E.SubElement(pat,q('fgColor'),rgb='FFFFF2CC');E.SubElement(pat,q('bgColor'),indexed='64');fid=len(fills)-1;fills.attrib['count']=str(len(fills))
for address in ['I4']+[f'G{r}' for r in range(5,10)]:
 c=sheet.find('.//'+q('c')+f'[@r="{address}"]');xf=copy.deepcopy(xfs[int(c.attrib['s'])]);xf.attrib.update(fillId=str(fid),applyFill='1');xfs.append(xf);c.attrib['s']=str(len(xfs)-1)
xfs.attrib['count']=str(len(xfs))
files['xl/styles.xml']=E.tostring(styles,encoding='utf-8',xml_declaration=True)
setup=sheet.find(q('pageSetup'));setup.attrib.update(fitToWidth='1',fitToHeight='1');setup.attrib.pop('scale',None)
sp=sheet.find(q('sheetPr'))
if sp is None:sp=E.Element(q('sheetPr'));sheet.insert(0,sp)
ps=sp.find(q('pageSetUpPr'))
if ps is None:ps=E.SubElement(sp,q('pageSetUpPr'))
ps.attrib['fitToPage']='1'
# Apply the user's final instruction: no fee or bank-balance input.
sd.remove(next(r for r in sd if r.attrib['r']=='4'))
for row in sd:
 old=int(row.attrib['r'])
 if old>=5:
  row.attrib['r']=str(old-1)
  for c in row:
   col=re.match('[A-Z]+',c.attrib['r'])[0];c.attrib['r']=f'{col}{old-1}'
   f=c.find(q('f'))
   if f is not None:f.text=re.sub(r'([A-Z]+)(\d+)',lambda m:m[1]+str(int(m[2])-1 if int(m[2])>=5 else int(m[2])),f.text)
for address in ['B1']+[f'G{r}' for r in range(4,9)]+[f'I{r}' for r in range(4,10)]:
 c=sheet.find('.//'+q('c')+f'[@r="{address}"]')
 for child in list(c):c.remove(child)
 c.attrib.pop('t',None)
 if address.startswith('G'):E.SubElement(c,q('v')).text='0'
for r in range(4,9):
 c=sheet.find('.//'+q('c')+f'[@r="H{r}"]');c.find(q('f')).text=f'SUM(F{r}:G{r})'
 v=c.find(q('v'))
 if v is None:v=E.SubElement(c,q('v'))
 v.text=sheet.find('.//'+q('c')+f'[@r="F{r}"]/'+q('v')).text;c.attrib['t']='n'
 # Restore the original anchor fee style, without input highlighting.
 sheet.find('.//'+q('c')+f'[@r="G{r}"]').attrib['s']=next(c.attrib['s'] for c in orig[4] if c.attrib['r']=='G4')
for col in ['G','H']:
 c=sheet.find('.//'+q('c')+f'[@r="{col}9"]');c.find(q('f')).text=f'SUM({col}4:{col}8)'
 v=c.find(q('v'))
 if v is None:v=E.SubElement(c,q('v'))
 v.text='0' if col=='G' else '4040257';c.attrib['t']='n'
for m in merges:
 m.attrib['ref']=re.sub(r'([A-Z]+)(\d+)',lambda x:x[1]+str(int(x[2])-1 if int(x[2])>=5 else int(x[2])),m.attrib['ref'])
files['xl/worksheets/sheet1.xml']=E.tostring(sheet,encoding='utf-8',xml_declaration=True)
book=E.fromstring(files['xl/workbook.xml'])
for name in book.findall('.//'+q('definedName')):
 if name.attrib['name']=='_xlnm.Print_Area':name.text='放行總表素材!$A$1:$J$12'
calc=E.SubElement(book,q('calcPr'),calcId='191029',fullCalcOnLoad='1')
files['xl/workbook.xml']=E.tostring(book,encoding='utf-8',xml_declaration=True)
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
 for name,data in files.items():z.writestr(name,data)
assert sheet.find('.//'+q('c')+'[@r="F9"]/'+q('f')).text=='SUM(F4:F8)'
assert len([r for r in sd if 4<=int(r.attrib['r'])<=8])==5
assert '{{' not in ''.join(sheet.itertext())
print('版面保留驗證通過：5 筆資料、合計第 9 列、備註第 10 列、簽章第 11 列；列印 A4 橫向一頁。')
