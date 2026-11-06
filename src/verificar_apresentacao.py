"""Conferência da apresentação contra a fonte e entre formatos."""
import json,hashlib
from pathlib import Path
from decimal import Decimal
from openpyxl import load_workbook
from docx import Document
from pypdf import PdfReader
from apresentacao import display,cell_display

def main():
    root=Path(__file__).resolve().parents[1];cfg=json.loads((root/'config.local.json').read_text(encoding='utf-8-sig'));out=root/cfg['saida'];run=out/(out/'ULTIMA_EXECUCAO.txt').read_text().strip()
    plans=json.loads((run/'dados/apresentacao.json').read_text(encoding='utf-8'))['quadros'];proof=json.loads((run/'dados/conferencia_apresentacao.json').read_text(encoding='utf-8'));raw=json.loads((run/'dados/transcricao_fonte.json').read_text(encoding='utf-8'))
    assert hashlib.sha256((root/cfg['arquivo']).read_bytes()).hexdigest()==proof['sha256_fonte']
    source=load_workbook(root/cfg['arquivo'],data_only=True)[cfg['aba']]
    for item in raw:
        i=item['Linha na fonte'];assert item['Servidor']==source[f'B{i}'].value
        for label,col in [('Valor adiantado','E'),('Valor utilizado','F'),('Valor devolvido','G')]:assert Decimal(str(item[label])).quantize(Decimal('.01'))==Decimal(str(source[f'{col}{i}'].value)).quantize(Decimal('.01'))
    book=load_workbook(run/'Planilhas/Analise_editavel.xlsx',data_only=False)
    doc=Document(run/'Relatorio para o RCI editavel.docx');assert len(doc.tables)==len(plans)
    for plan,table in zip(plans,doc.tables):
        ws=book[plan['id'][:31]]
        assert [c.text for c in table.rows[0].cells]==plan['cabecalhos']
        assert [ws.cell(5,j+1).value for j in range(len(plan['cabecalhos']))]==plan['cabecalhos']
        assert sum(plan['larguras'])==480 and len(plan['cabecalhos'])<=6
        for i,row in enumerate(plan['linhas']):
            for j,value in enumerate(row):
                assert ws.cell(i+6,j+1).value==value,(plan['id'],i,j,'Excel divergente')
                assert table.rows[i+1].cells[j].text==cell_display(plan,i,j),(plan['id'],i,j,'Word divergente')
        assert not any('_' in x for x in plan['cabecalhos'])
        assert Path(str(run/plan['imagem'])+'.png').exists()
    pdf=PdfReader(run/'Relatorio para o RCI.pdf');assert len(pdf.pages)==len(plans)
    for page,plan in zip(pdf.pages,plans):assert plan['titulo'] in page.extract_text()
    for row in book['Mensal'].iter_rows(min_row=6,max_row=11,max_col=1):assert '/' in row[0].value and '-' not in row[0].value
    result={'resultado':'Conferência aprovada','linhas_fonte':len(raw),'quadros':len(plans),'valores_nomes_conferidos_com_fonte':True,'excel_word_mesmos_dados':True,'cabecalhos_sem_codigos':True,'largura_quadros_pt':480,'pdf_paginas':len(pdf.pages),'layout_word':'Exportação e verificação visual necessárias'}
    print(json.dumps(result,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
