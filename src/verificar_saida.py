"""Verifica integridade, reconciliação e tipos dos artefatos de uma execução."""
import argparse,json,hashlib,csv
from pathlib import Path
from collections import Counter
from decimal import Decimal
from openpyxl import load_workbook
from docx import Document
from pypdf import PdfReader

def main():
    p=argparse.ArgumentParser();p.add_argument('--config',default='config.local.json');args=p.parse_args()
    root=Path(__file__).resolve().parents[1];cfg=json.loads((root/args.config).read_text(encoding='utf-8-sig'));out=root/cfg['saida'];run=out/(out/'ULTIMA_EXECUCAO.txt').read_text().strip()
    manifest=json.loads((run/'manifesto.json').read_text(encoding='utf-8'));s=json.loads((run/'dados/indicadores.json').read_text(encoding='utf-8'));rows=json.loads((run/'dados/base_normalizada.json').read_text(encoding='utf-8'))
    assert hashlib.sha256((root/cfg['arquivo']).read_bytes()).hexdigest()==manifest['sha256'],'Fonte mudou: executar novamente'
    assert len(rows)==s['registros'];assert len({r['processo'] for r in rows})==s['processos_unicos']
    for field in ['adiantado','utilizado','devolvido']:
        assert sum(Decimal(str(r[field])) for r in rows if r[field] is not None)==Decimal(str(s[field])),field
    for field in ['smf','cgm']:
        counts=Counter(r[field] for r in rows);assert sum(counts.values())==s['registros']
    w=load_workbook(run/'Planilhas/Analise_editavel.xlsx',data_only=False)
    assert len(w['Base normalizada'].tables)==1 or w['Base normalizada'].auto_filter.ref
    assert w['Resumo']['B5'].data_type=='n';assert w['Resumo']['B9'].data_type=='n'
    for ws in w:
        for cells in ws:
            for cell in cells:assert cell.data_type!='e',f'Erro de Excel: {ws.title}!{cell.coordinate}'
    pages={}
    for path in run.glob('*.pdf'):
        pdf=PdfReader(path);pages[path.name]=len(pdf.pages)
        for page in pdf.pages:assert len(page.extract_text() or '')>50,'Página vazia'
    for path in run.glob('*.docx'):
        d=Document(path);assert len(d.paragraphs)>5;assert all(p.text is not None for p in d.paragraphs)
    for file in manifest['arquivos']:
        assert hashlib.sha256((run/file['caminho']).read_bytes()).hexdigest()==file['sha256'],file['caminho']
    print(json.dumps({'verificacao':'PASSOU','registros':len(rows),'processos_unicos':s['processos_unicos'],'abas_xlsx':len(w.sheetnames),'pdf_paginas':pages,'hash_fonte_confere':True,'hashes_arquivos_conferem':True,'layout_docx':'Não renderizado; validação estrutural apenas'},ensure_ascii=False,indent=2))

if __name__=='__main__':main()
