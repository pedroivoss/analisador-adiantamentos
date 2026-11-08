"""Conferência exaustiva dos campos literais e da conciliação mensal."""
from pathlib import Path
from hashlib import sha256
import json
from openpyxl import load_workbook
from docx import Document
from analisador import cents
from apresentacao import moeda

def main():
    root=Path(__file__).resolve().parents[1];cfg=json.loads((root/'config.local.json').read_text(encoding='utf-8-sig'));out=root/cfg['saida'];run=out/(out/'ULTIMA_EXECUCAO.txt').read_text().strip()
    p=json.loads((run/'dados/complementos.json').read_text(encoding='utf-8'));assert sha256(Path(p['fonte']).read_bytes()).hexdigest()==p['sha256']
    src=load_workbook(p['fonte'],data_only=True);ws=src[p['aba_detalhe']];con=src[p['aba_consolidado']]
    book=load_workbook(run/'Planilhas/Analise mensal e detalhada.xlsx',data_only=True)
    for i,(row,trace) in enumerate(zip(p['detalhe'],p['rastreio']),6):
        assert row==[ws[f'{c}{trace["linha_fonte"]}'].value for c in p['colunas_fonte']]
        assert row==[book['Relação individual'].cell(i,j).value for j in range(1,14)],f'Campo literal alterado na linha {i}'
    for m in p['meses']:
        sheet=book[m['rotulo'].replace('/',' ')];base=3+(int(m['chave'][-2:])-1)*3
        for i,row in enumerate(m['linhas'],6):
            assert row==[sheet.cell(i,j).value for j in range(1,6)]
            si=i-3;assert row[:2]==[con.cell(si,j).value for j in [1,2]]
            assert [cents(v) for v in row[2:]]==[cents(con.cell(si,j).value) for j in range(base,base+3)]
        assert [cents(sheet.cell(27,j).value) for j in [3,4,5]]==[cents(v) for v in m['totais']]
    for i,row in enumerate(p['presenca'],6):assert row==[book['Presença mensal'].cell(i,j).value for j in range(1,9)]
    d=Document(run/'Relacao individual dos adiantamentos.docx');actual=[[c.text for c in row.cells] for table in d.tables for row in table.rows[1:]]
    expected=[[moeda(v) if j in [3,4,5] else ('' if v is None else str(v)) for j,v in enumerate(row)] for row in p['detalhe']]
    assert actual==expected,'Relação no Word difere da fonte.'
    monthly=Document(run/'Analise mensal por secretaria.docx');assert len(monthly.tables)==7
    for table,m in zip(monthly.tables,p['meses']):
        expected=[[r[0],r[1],*[moeda(v) for v in r[2:]]] for r in m['linhas']]+[['Total','',*[moeda(v) for v in m['totais']]]]
        assert [[c.text for c in row.cells] for row in table.rows[1:]]==expected
    result={'fonte_sha256':p['sha256'],'registros':len(p['detalhe']),'campos_literais_conferidos':len(p['detalhe'])*13,'valores_mensais_conciliados':len(p['conferencias']),'totais_mensais_conferidos':18,'secretarias_consolidadas':len(p['presenca']),'tabelas_word_conferidas':len(monthly.tables)+len(d.tables),'resultado':'Aprovado'}
    (run/'dados/validacao_complementos.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=False,indent=2))
if __name__=='__main__':main()
