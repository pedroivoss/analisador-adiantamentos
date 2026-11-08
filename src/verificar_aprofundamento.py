"""Conferências de fidelidade do texto e dos quadros preservados no relatório."""
from pathlib import Path
from hashlib import sha256
from zipfile import ZipFile
import json,re
from docx import Document
from pypdf import PdfReader
from apresentacao import cell_display

def main():
    root=Path(__file__).resolve().parents[1];cfg=json.loads((root/'config.local.json').read_text(encoding='utf-8-sig'));out=root/cfg['saida'];run=out/(out/'ULTIMA_EXECUCAO.txt').read_text().strip()
    analysis=json.loads((run/'dados/analise_aprofundada.json').read_text(encoding='utf-8'))
    evidence=json.loads((run/'dados/validacao_aprofundamento.json').read_text(encoding='utf-8'))
    plans=json.loads((run/'dados/apresentacao.json').read_text(encoding='utf-8'))['quadros']
    assert sha256((root/cfg['arquivo']).read_bytes()).hexdigest()==evidence['sha256_fonte']
    assert all(evidence['quadros_graficos_preservados'].values())
    doc=Document(run/'Relatorio completo para o RCI.docx')
    selected=[plans[i] for i in evidence['layout']['indices_tabelas_base']]
    assert len(doc.tables)==len(selected)+(7 if (run/'dados/complementos.json').exists() else 0)
    for table,plan in zip(doc.tables,selected):
        assert [c.text for c in table.rows[0].cells]==plan['cabecalhos'],plan['id']
        for i,row in enumerate(plan['linhas']):
            assert [c.text for c in table.rows[i+1].cells]==[cell_display(plan,i,j) for j in range(len(row))],plan['id']
    public='\n'.join(p.text for p in doc.paragraphs)
    for prohibited in ['Critério.','valor adt','valor utl','valor dev','mes_informado','linha_origem','coluna AB']:
        assert not re.search(r'\b'+re.escape(prohibited)+r'(?!\w)',public),'Instrução operacional vazou para o texto do RCI: '+prohibited
    assert len(doc.inline_shapes)==5,'Gráficos faltantes ou inesperados'
    for plan in selected:
        for paragraph in analysis['textos'][plan['id']]:assert paragraph in public
    originals={(o['processo'],o['nota']):o for o in json.loads((run/'dados/observacoes_consultadas.json').read_text(encoding='utf-8'))}
    for case in analysis['casos']:
        obs=originals[(case['processo'],case['nota'])];assert sha256(obs['texto'].encode()).hexdigest()==case['sha256_observacao']
    pdf=PdfReader(run/'Relatorio completo para o RCI.pdf');assert len(pdf.pages)>0
    for page in pdf.pages:assert len(page.extract_text() or '')>30,'Página vazia'
    print(json.dumps({'resultado':'Conferência aprovada','tabelas_preservadas':len(selected),'graficos_preservados':len(doc.inline_shapes),'casos_com_evidencia_revisada':len(analysis['casos']),'paginas_pdf':len(pdf.pages),'palavras_analiticas':evidence['layout']['palavras_analise'],'sha256_fonte_confere':True,'sem_criterios_operacionais_no_corpo':True},ensure_ascii=False,indent=2))

if __name__=='__main__':main()
