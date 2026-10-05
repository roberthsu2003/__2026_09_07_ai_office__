from pathlib import Path
import zipfile,copy,re
from lxml import etree as E
import openpyxl
N='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/package/2006/relationships'
C='http://schemas.openxmlformats.org/package/2006/content-types'
q=lambda t:'{'+N+'}'+t
p=Path('outputs/01a10bc7-8ae6-7fc0-82fa-7915b40e0642/出納付款放行總表_已完成.xlsx')
with zipfile.ZipFile(p) as z: a=E.fromstring(z.read('xl/worksheets/sheet1.xml'))
with zipfile.ZipFile('素材_放行總表.xlsx') as z: files={i.filename:z.read(i.filename) for i in z.infolist()}
s=E.fromstring(files['xl/worksheets/sheet1.xml'])
orig={int(r.get('r')):r for r in s.find(q('sheetData'))}
sd=s.find(q('sheetData'))
for r in list(sd):sd.remove(r)
for ar in a.find(q('sheetData')):
 r=int(ar.get('r'));src=r if r<=3 or r>=9 else 4
 row=copy.deepcopy(orig.get(src,orig[12]))
 attrs={k:v for k,v in ar.attrib.items() if not k.startswith('{')}
 row.attrib.clear();row.attrib.update(attrs)
 for c in list(row):row.remove(c)
 stylemap={re.match('[A-Z]+',c.get('r'))[0]:c.get('s') for c in orig.get(src,orig[12])}
 for ac in ar:
  c=copy.deepcopy(ac);col=re.match('[A-Z]+',c.get('r'))[0];style=stylemap.get(col,'28' if r==10 else None)
  if style:c.set('s',style)
  else:c.attrib.pop('s',None)
  row.append(c)
 sd.append(row)
for tag in ['mergeCells','pageSetup','sheetPr']:
 old=s.find(q(tag));new=a.find(q(tag))
 if old is not None and new is not None:s.replace(old,copy.deepcopy(new))
 elif new is not None:s.insert(0,copy.deepcopy(new))
# Native printer relationship is retained with the source's printer settings.
s.find(q('pageSetup')).set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id','rId1')
dimension=s.find(q('dimension'))
if dimension is not None:dimension.set('ref','A1:J12')
book=E.fromstring(files['xl/workbook.xml'])
external=book.find(q('externalReferences'))
if external is not None:book.remove(external)
names=book.find(q('definedNames'))
for name in list(names):
 if name.get('name')=='銀行名稱':names.remove(name)
calc=book.find(q('calcPr'));calc.set('fullCalcOnLoad','1')
rels=E.fromstring(files['xl/_rels/workbook.xml.rels'])
for rel in list(rels):
 if rel.get('Type').endswith(('/externalLink','/calcChain')):rels.remove(rel)
types=E.fromstring(files['[Content_Types].xml'])
for item in list(types):
 if item.get('PartName','').startswith('/xl/externalLinks/') or item.get('PartName')=='/xl/calcChain.xml':types.remove(item)
for name,xml in [('xl/worksheets/sheet1.xml',s),('xl/workbook.xml',book),('xl/_rels/workbook.xml.rels',rels),('[Content_Types].xml',types)]:files[name]=E.tostring(xml,encoding='UTF-8',xml_declaration=True,standalone=True)
files={k:v for k,v in files.items() if not k.startswith('xl/externalLinks/') and k!='xl/calcChain.xml'}
with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
 for name,data in files.items():z.writestr(name,data)
w=openpyxl.load_workbook(p,data_only=False)
assert w.active['F9'].value=='=SUM(F4:F8)'
assert sum(w.active[f'F{r}'].value for r in range(4,9))==4040257
assert w.active['A10'].value=='本期手續費由公司負擔；單筆大額款項請主管覆核放行。'
with zipfile.ZipFile(p) as z:
 assert z.testzip() is None
 for name in z.namelist():
  if name.endswith('.xml') or name.endswith('.rels'):E.fromstring(z.read(name))
print('已使用原始 Excel 原生封裝與樣式修復，通過獨立讀取、公式、金額與 XML 驗證。')
