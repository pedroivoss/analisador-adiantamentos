"""Configuração OOXML de impressão, recurso ausente na API de autoria usada."""
from zipfile import ZipFile,ZIP_DEFLATED
from pathlib import Path
from xml.etree import ElementTree as ET
import os
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
def configure_print(path,records):
    path=Path(path);ET.register_namespace('',NS);q=lambda x:'{'+NS+'}'+x
    with ZipFile(path) as z:parts={n:z.read(n) for n in z.namelist()}
    book=ET.fromstring(parts['xl/workbook.xml']);defs=book.find(q('definedNames'))
    if defs is None:defs=ET.SubElement(book,q('definedNames'))
    for i,s in enumerate(book.find(q('sheets'))):
        name=s.attrib['name'];landscape=i==7;end='M' if landscape else ('H' if i==6 else 'E');last=records+6 if landscape else (26 if i==6 else 27)
        ref="'"+name.replace("'","''")+"'!"
        for kind,target in [('_xlnm.Print_Area',ref+f'$A$2:${end}${last}'),('_xlnm.Print_Titles',ref+'$1:$5')]:
            el=ET.SubElement(defs,q('definedName'),{'name':kind,'localSheetId':str(i)});el.text=target
        key=f'xl/worksheets/sheet{i+1}.xml';sheet=ET.fromstring(parts[key]);pr=sheet.find(q('sheetPr'))
        if pr is None:pr=ET.Element(q('sheetPr'));sheet.insert(0,pr)
        props=pr.find(q('pageSetUpPr'))
        if props is None:props=ET.SubElement(pr,q('pageSetUpPr'))
        props.set('fitToPage','1')
        for tag in ['pageMargins','pageSetup']:
            old=sheet.find(q(tag))
            if old is not None:sheet.remove(old)
        ET.SubElement(sheet,q('pageMargins'),{'left':'0.4','right':'0.4','top':'0.45','bottom':'0.45','header':'0.2','footer':'0.2'})
        ET.SubElement(sheet,q('pageSetup'),{'paperSize':'1','orientation':'landscape' if landscape else 'portrait','fitToWidth':'1','fitToHeight':'0' if landscape else '1'})
        parts[key]=ET.tostring(sheet,encoding='utf-8',xml_declaration=True)
    parts['xl/workbook.xml']=ET.tostring(book,encoding='utf-8',xml_declaration=True)
    temp=path.with_suffix('.print.xlsx')
    with ZipFile(temp,'w',ZIP_DEFLATED) as z:
        for n,b in parts.items():z.writestr(n,b)
    os.replace(temp,path)
