import zipfile, xml.etree.ElementTree as E
from pathlib import Path
path=Path('outputs/01a10bc7-8ae6-7fc0-82fa-7915b40e0642/放行總表_佔位字符樣版.xlsx')
ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
with zipfile.ZipFile('素材_放行總表.xlsx') as source, zipfile.ZipFile(path) as result:
    files={i.filename:result.read(i.filename) for i in result.infolist()}
    orig=E.fromstring(source.read('xl/worksheets/sheet1.xml'))
    sheet=E.fromstring(files['xl/worksheets/sheet1.xml'])
    for tag in ['printOptions','pageMargins','pageSetup','headerFooter']:
        item=orig.find('{'+ns+'}'+tag)
        if item is not None:
            item.attrib.pop('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id',None)
            sheet.append(item)
    files['xl/worksheets/sheet1.xml']=E.tostring(sheet,encoding='utf-8',xml_declaration=True)
    book=E.fromstring(files['xl/workbook.xml'])
    names=book.find('{'+ns+'}definedNames')
    for item in E.fromstring(source.read('xl/workbook.xml')).find('{'+ns+'}definedNames'):
        if item.attrib['name'].startswith('_xlnm.Print_'):
            names.append(item)
    files['xl/workbook.xml']=E.tostring(book,encoding='utf-8',xml_declaration=True)
with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as result:
    for name,data in files.items(): result.writestr(name,data)
print('已保留原表列印範圍、橫向 A4、頁首頁尾與邊界。')
