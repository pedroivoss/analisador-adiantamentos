"""Relatório substantivo de adiantamentos com as tabelas aprovadas preservadas."""
from pathlib import Path
from datetime import datetime,date
from collections import Counter
from copy import deepcopy
import sys,json,hashlib,shutil,argparse
from analisador import analyze,json_write,csv_write,load_workbook
from apresentacao import cell_display
from textos_analiticos import build,reais,pct

FEDERAL_SOURCE='https://www.gov.br/compras/pt-br/acesso-a-informacao/comunicados/2025/no-47-25-decreto-altera-valores-da-lei-14-133-para-compras-publicas'

def norm_matrix():
    return [
      ('Finalidade e excepcionalidade','IN, arts. 1º, 3º, §2º, 4º, IV, e 14, §2º; Manual, pp. 7–11.','Confrontar as justificativas e o objeto. As observações registradas permitem contextualizar casos; não demonstram a urgência de todos os pedidos.'),
      ('Limite da concessão','IN, art. 3º, retificação de 2025; Decreto federal nº 12.807/2025.','A redação atual vincula o limite ao valor atualizado da dispensa e prevê hipótese diferenciada de viagem. Não aplicar automaticamente o valor fixo de R$ 4.000,00 do Manual antigo.'),
      ('Natureza e objeto','IN, arts. 4º, III e §1º, e 9º; Manual, p. 12.','Natureza registrada e relatos de incompatibilidade são analisáveis. A verificação item a item exige comprovantes e autorização original.'),
      ('Período de aplicação','IN, arts. 6º e 8º; Manual, p. 11.','Pagamento, última aplicação e prazo permitem diferenças de datas. A última aplicação não prova a data da primeira despesa nem a regularidade de todas as notas fiscais.'),
      ('Modalidade de pagamento','IN, art. 8º, parágrafo único; Manual, pp. 11, 18 e 21.','Observações registram exemplos de crédito. Não há campo estruturado que permita contar todos os pagamentos por modalidade.'),
      ('Identificação e integridade fiscal','IN, art. 10; Manual, pp. 13–15.','Validade fiscal, CNPJ, discriminação e legibilidade exigem os documentos. Exemplos relatados são atribuídos às observações, sem nova validação fiscal.'),
      ('Justificação e relatório circunstanciado','IN, arts. 11, 12 e 19, II; Manual, pp. 9, 13 e 15.','Relatos de ausência ou insuficiência são contextualizados. Não se presume inexistência de documentação quando falta observação.'),
      ('Modicidade de refeições','IN, art. 16; Manual, pp. 11–12.','Não apurável de forma integral: faltam o valor por refeição, a quantidade de beneficiários, sua condição e as justificativas excepcionais.'),
      ('Devolução do saldo','IN, art. 13; Manual, pp. 15, 18 e 21.','Valor, data, vencimento e atraso informado são disponíveis. Calendário municipal e comprovantes de depósito devem ser validados; não há extrato bancário para confirmação do ingresso.'),
      ('Prestação e protocolo','IN, art. 17 e §§1º–3º; Manual, p. 13.','Data e atraso informados são disponíveis. Confirmar se o marco é início de elaboração ou protocolo na SMF; cancelamento pode alterar o marco aplicável.'),
      ('Encerramento do exercício','IN, art. 18 e art. 20, §6º.','O recorte é de pagamentos do primeiro semestre. Não aplicar automaticamente a regra anual sem verificar o decreto específico de encerramento do exercício.'),
      ('Organização e documentos obrigatórios','IN, art. 19 e §§1º–3º; Manual, pp. 13–17.','Relatos permitem identificar exemplos. Orientações de SEI mencionadas nas observações não foram presumidas como integralmente lidas.'),
      ('Exame SMF e saneamento','IN, arts. 20 e 21; Manual, pp. 16–17.','Categorias e ausências podem ser consolidadas. Prazos de análise, notificações e resposta não podem ser medidos sem suas datas.'),
      ('Revisão pela CGM','IN, art. 22 e parágrafo único.','Categorias permitem comparação. Não há datas suficientes para medir o prazo de análise da CGM ou o retorno das recomendações.'),
      ('Impedimentos para concessão','IN, arts. 2º e 14.','Frequência de registros não prova simultaneidade de contas abertas, sequência temporal de ressalvas fazendárias ou ciência de rejeição antes de nova concessão.'),
      ('Ressarcimento e responsabilização','IN, arts. 24–29; Manual, pp. 18–21.','Montantes de contas rejeitadas não equivalem a dano ou glosa definitiva. Exigem identificação de despesas impugnadas, ato fazendário, notificação e desfecho.'),
    ]

def render_reports(root,run,plans,analysis,meta,template):
    from docx import Document
    from docx.shared import Pt,Inches,RGBColor
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from reportlab.platypus import SimpleDocTemplate,Paragraph,Table,TableStyle,Spacer,PageBreak,Image,KeepTogether
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader
    from xml.sax.saxutils import escape
    original=Document(template)
    native_tables=[deepcopy(t._tbl) for t in original.tables]
    assert len(native_tables)==len(plans),'Número de tabelas do modelo difere dos quadros aprovados.'
    doc=Document(template)
    body=doc._element.body
    for child in list(body):
        if child.tag!=qn('w:sectPr'):body.remove(child)
    for sty in ['Normal','Title','Heading 1','Heading 2']:
        doc.styles[sty].font.name='Arial';doc.styles[sty].font.color.rgb=RGBColor(0,0,0)
        for border in doc.styles[sty].element.findall('.//'+qn('w:pBdr')):border.getparent().remove(border)
    doc.styles['Normal'].font.size=Pt(10.5);doc.styles['Normal'].paragraph_format.line_spacing=1.08;doc.styles['Normal'].paragraph_format.space_after=Pt(7)
    doc.styles['Heading 1'].font.size=Pt(14);doc.styles['Heading 1'].font.bold=True
    doc.styles['Heading 2'].font.size=Pt(11);doc.styles['Heading 2'].font.bold=True
    doc.core_properties.title='Análise dos adiantamentos do primeiro semestre de 2026';doc.core_properties.author=''
    pstyle=ParagraphStyle('Corpo',fontName='Helvetica',fontSize=10.2,leading=13.4,spaceAfter=8)
    hstyle=ParagraphStyle('Secao',fontName='Helvetica-Bold',fontSize=14,leading=17,spaceAfter=10)
    title=ParagraphStyle('Titulo',fontName='Helvetica-Bold',fontSize=17,leading=21,spaceAfter=12)
    note=ParagraphStyle('Fonte',fontName='Helvetica',fontSize=8,leading=10,spaceAfter=7)
    cell=ParagraphStyle('Celula',fontName='Helvetica',fontSize=8.4,leading=10.5)
    header=ParagraphStyle('Cabecalho',parent=cell,fontName='Helvetica-Bold',textColor=colors.white)
    story=[];public_text=[]
    def pdfp(text,style=pstyle):return Paragraph(escape(text),style)
    def para(text):
        doc.add_paragraph(text);story.append(pdfp(text));public_text.append(text)
    def heading(text,cover=False):
        doc.add_paragraph(text,'Title' if cover else 'Heading 1');story.append(pdfp(text,title if cover else hstyle));public_text.append(text)
    def pagebreak():doc.add_page_break();story.append(PageBreak())
    def pdf_table(plan):
        data=[[pdfp(x,header) for x in plan['cabecalhos']]]
        data.extend([[pdfp(cell_display(plan,i,j),cell) for j in range(len(row))] for i,row in enumerate(plan['linhas'])])
        t=Table(data,colWidths=plan['larguras'],repeatRows=1)
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#24476B')),('GRID',(0,0),(-1,-1),.35,colors.HexColor('#D9D9D9')),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F2F5F8')]),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5)]))
        return t
    source='Fonte: controle de prestações de contas de adiantamentos do Município de Ubatuba.'
    heading('Análise dos adiantamentos do primeiro semestre de 2026',True)
    para('Análise para revisão da Controladoria Geral do Município')
    for text in analysis['introducao']:para(text)
    heading('Enquadramento da análise')
    para('A apreciação considera a finalidade e a excepcionalidade do regime, a compatibilidade da natureza das despesas, os prazos de aplicação e encerramento e a suficiência da comprovação. As referências centrais são a IN-001/2018-CGM, na versão retificada em 2025, e o Manual de Adiantamentos. A competência e o fluxo dos exames são tratados conforme os arts. 20 a 22 da IN. As medidas financeiras eventualmente decorrentes de impugnação devem observar os arts. 24 e 25.')
    main_indices=[]
    for i,plan in enumerate(plans):
        if plan['id']=='Revisão dos dados':continue
        pagebreak();heading(plan['titulo']);main_indices.append(i)
        native=deepcopy(native_tables[i])
        # Mantém o quadro compacto inteiro, com linhas indivisíveis.
        trs=native.findall(qn('w:tr'))
        for ri,tr in enumerate(trs):
            for p in tr.findall('.//'+qn('w:p')):
                pp=p.find(qn('w:pPr'))
                if pp is None:pp=OxmlElement('w:pPr');p.insert(0,pp)
                for old in pp.findall(qn('w:keepNext')):pp.remove(old)
                if ri<len(trs)-1:pp.append(OxmlElement('w:keepNext'))
        doc._element.body.insert(len(doc._element.body)-1,native)
        story.extend([KeepTogether([pdf_table(plan),pdfp(source,note)]),Spacer(1,5)])
        pp=doc.add_paragraph(source)
        for rr in pp.runs:rr.font.size=Pt(8)
        chart={'Mensal':'mensal','Secretarias':'secretarias','Naturezas':'naturezas','Durações':'prazos','Exame CGM':'exames'}.get(plan['id'])
        if chart:
            pic=run/'graficos'/(chart+'.png');iw,ih=ImageReader(str(pic)).getSize();w=480;h=ih*w/iw
            # Dimensões naturais e legenda vinculada: a figura nunca se divide.
            limit=165 if plan['id']=='Mensal' else 300
            if h>limit:w*=limit/h;h=limit
            doc.add_picture(str(pic),width=Inches(w/72))
            story.extend([KeepTogether([Image(str(pic),width=w,height=h),pdfp(source,note)]),Spacer(1,7)])
        for text in analysis['textos'][plan['id']]:para(text)
        part=run/'Textos analíticos';part.mkdir(exist_ok=True)
        (part/(plan['id']+'.txt')).write_text(plan['titulo']+'\n\n'+'\n\n'.join(analysis['textos'][plan['id']]),encoding='utf-8')
    pagebreak();heading('Principais apontamentos e implicações para o controle')
    for text in analysis['achados']:para(text)
    if analysis['casos']:
        heading('Casos registrados nas contas com rejeição na CGM')
        for case in analysis['casos']:
            doc.add_paragraph('Nota '+case['nota']+' — '+case['secretaria'],'Heading 2');story.append(pdfp('Nota '+case['nota']+' — '+case['secretaria'],hstyle))
            para(case['sintese'])
            doc.paragraphs[-1].paragraph_format.keep_with_next=True
            # A referência do processo fica no caderno de conferência; o texto público não reproduz instruções operacionais.
            para('Referência de análise: '+case['referencia'])
    pagebreak();heading('Conclusão e recomendações')
    for text in analysis['conclusao']:para(text)
    heading('Referências')
    for text in analysis['referencias'].values():para(text)
    doc.save(run/'Relatorio completo para o RCI.docx')
    def footer(c,d):
        c.setFont('Helvetica',8);c.drawString(66,24,'Adiantamentos - primeiro semestre de 2026');c.drawRightString(546,24,str(d.page))
    SimpleDocTemplate(str(run/'Relatorio completo para o RCI.pdf'),pagesize=(612,792),leftMargin=66,rightMargin=66,topMargin=42,bottomMargin=42,title='Análise dos adiantamentos do primeiro semestre de 2026',author='').build(story,onFirstPage=footer,onLaterPages=footer)
    (run/'Texto integral para o RCI.txt').write_text('\n\n'.join(public_text),encoding='utf-8')
    return {'tabelas_no_relatorio':len(main_indices),'indices_tabelas_base':main_indices,'palavras_analise':sum(len(x.split()) for x in public_text),'tabelas_preservadas':True,'graficos_incluidos':5}

def render_caderno(root,run,analysis,rows,tables,meta,observations):
    from docx import Document
    from docx.shared import Inches,Pt,RGBColor
    from reportlab.platypus import SimpleDocTemplate,Paragraph,Table,TableStyle,Spacer,PageBreak
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from xml.sax.saxutils import escape
    doc=Document();sec=doc.sections[0];sec.page_width=Inches(8.5);sec.page_height=Inches(11);sec.left_margin=sec.right_margin=Inches(66/72)
    for name in ['Normal','Title','Heading 1','Heading 2']:doc.styles[name].font.name='Arial';doc.styles[name].font.color.rgb=RGBColor(0,0,0)
    doc.styles['Normal'].font.size=Pt(10)
    body=ParagraphStyle('Corpo',fontName='Helvetica',fontSize=9.5,leading=12.5,spaceAfter=7);head=ParagraphStyle('Titulo',fontName='Helvetica-Bold',fontSize=14,leading=18,spaceAfter=10)
    story=[]
    def p(text,title=False):
        doc.add_paragraph(text,'Heading 1' if title else 'Normal');story.append(Paragraph(escape(text),head if title else body))
    p('Caderno de conferência da análise de adiantamentos',True)
    p('Documento de apoio à revisão. Os textos operacionais, critérios de extração e pendências abaixo foram separados do corpo destinado ao RCI.')
    p('Fonte: '+meta['arquivo']+'. Aba: '+meta['aba']+'. SHA-256: '+meta['sha256'])
    p('Método e população',True)
    p('Pagamentos entre '+meta['config']['inicio']+' e '+meta['config']['fim']+'. Os valores são calculados em centavos e conciliados com a fonte. Nomes semelhantes não são unificados. Os quadros e gráficos aprovados foram preservados; o corpo do relatório foi ampliado sem usar IA para produzir números.')
    p('Foram lidas as páginas 43 a 52 do RCI de referência, a IN fornecida e os capítulos 2 a 8 do Manual. O modelo de RCI contém dados de 2025 e não foi utilizado como população de comparação. O relatório não conclui redução ou aumento de irregularidades entre períodos de cobertura diferente.')
    p('As observações da coluna AB foram lidas para contextualização. Colunas auxiliares e rascunhos fora do quadro principal não foram usados. Não se contou prevalência de causas por simples busca de palavras. As nove sínteses de rejeição foram revisadas e vinculadas a processo, nota e hash do texto observado; alterações futuras exigem nova revisão.')
    p('Cobertura normativa',True)
    for topic,ref,scope in norm_matrix():p(topic,True);p(ref);p(scope)
    cfg=meta['config'];year=date.fromisoformat(cfg['inicio']).year
    if year==2026:
        cap=654921;double=1309842;above=[r for r in rows if r['adiantado']>cap];above_double=[r for r in rows if r['adiantado']>double]
        p('Referência monetária de 2026',True)
        p('O Portal de Compras do Governo Federal informa R$ 65.492,11 para o art. 75, II, a partir de 1º de janeiro de 2026, conforme o Decreto nº 12.807/2025. A referência de 10% prevista na IN corresponde a R$ 6.549,21 em centavos; o dobro corresponde a R$ 13.098,42. Fonte oficial consultada em 06/11/2026: '+FEDERAL_SOURCE)
        p(f'Há {len(above)} registros acima da referência ordinária, somando {reais(sum(r["adiantado"] for r in above)/100)}, e {len(above_double)} acima da referência em dobro. Esses são sinalizadores para exame de enquadramento, autorização, tipo e duração da viagem e eventuais atos específicos. Não foram usados para alterar pareceres ou declarar automaticamente a irregularidade de concessões.')
        csv_write(run/'dados/conferir_limites_2026.csv',[{'Linha na fonte':r['linha_origem'],'Processo SEI':r['processo'],'Nota':r['nota'],'Servidor':r['servidor'],'Valor adiantado':r['adiantado']/100,'Acima da referência em dobro':r['adiantado']>double} for r in above])
    p('Pendências de consistência identificadas',True)
    for alert in tables['alertas']:
        if not alert['codigo'].startswith('PRAZO_DIVERGENTE'):p(f'Linha {alert["linha_origem"]}: {alert["descricao"]}')
    p('As referências de aplicação, devolução e prestação divergiram dos vencimentos registrados em diversas linhas. O calendário municipal não está validado. Os números de atrasos no corpo do RCI reproduzem os campos preenchidos; não se substituíram os vencimentos da fonte por uma contagem normativa automática.')
    p('Casos contextualizados',True)
    for case in analysis['casos']:
        p('Nota '+case['nota']+' — '+case['servidor'],True);p('Processo '+case['processo']+'; linha '+str(case['linha'])+'; secretaria '+case['secretaria']+'; SMF '+case['smf']+'; CGM '+case['cgm']+'.');p(case['sintese']);p(case['referencia'])
    for x in analysis['casos_alterados']:p('Síntese não reutilizada: '+str(x))
    p('Limites da conclusão',True)
    p('Não há documentos fiscais originais, extratos, datas de pareceres e de notificações suficientes para nova auditoria integral. Não se apuram dano, valores definitivos de glosa, impedimentos de novos adiantamentos ou tempestividade das análises SMF/CGM só com estes registros. A aplicação de regras locais deve considerar as versões normativas e autorizações efetivamente vigentes no caso concreto.')
    doc.save(run/'Caderno de conferencia.docx')
    SimpleDocTemplate(str(run/'Caderno de conferencia.pdf'),pagesize=(612,792),leftMargin=66,rightMargin=66,topMargin=42,bottomMargin=42).build(story)

def generate(root,base,cfg,fork=True):
    rows,tables,s,meta=analyze(cfg,root)
    old=json.loads((base/'manifesto.json').read_text(encoding='utf-8'))
    if old['sha256']!=meta['sha256']:raise RuntimeError('A fonte mudou depois da geração dos quadros. Rode analisar.ps1 para atualizar os quadros antes da análise textual.')
    if fork:
        run=base.parent/('ANALISE_COMPLETA_'+datetime.now().strftime('%Y%m%d-%H%M%S'))
        shutil.copytree(base,run,ignore=shutil.ignore_patterns('~$*'))
    else:run=base
    plans=json.loads((run/'dados/apresentacao.json').read_text(encoding='utf-8'))['quadros']
    ws=load_workbook(root/cfg['arquivo'],data_only=True)[cfg['aba']]
    observations=[]
    for r in rows:
        text=ws[f'AB{r["linha_origem"]}'].value
        observations.append({'linha':r['linha_origem'],'processo':r['processo'],'nota':r['nota'],'servidor_original':str(ws[f'B{r["linha_origem"]}'].value or ''),'secretaria':r['secretaria'],'smf':r['smf'],'cgm':r['cgm'],'texto':str(text) if text is not None else ''})
    analysis=build(root,rows,tables,s,meta,observations,plans)
    layout=render_reports(root,run,plans,analysis,meta,run/'Relatorio para o RCI editavel.docx')
    render_caderno(root,run,analysis,rows,tables,meta,observations)
    json_write(run/'dados/analise_aprofundada.json',analysis)
    json_write(run/'dados/observacoes_consultadas.json',observations)
    unchanged={}
    for p in [*base.glob('graficos/*'),*base.glob('Quadros para o RCI/*'),base/'Planilhas/Analise_editavel.xlsx']:
        if p.is_file():
            rel=p.relative_to(base);unchanged[str(rel)]=hashlib.sha256(p.read_bytes()).hexdigest()==hashlib.sha256((run/rel).read_bytes()).hexdigest()
    assert all(unchanged.values()),'Uma tabela ou gráfico foi alterado.'
    json_write(run/'dados/validacao_aprofundamento.json',{'sha256_fonte':meta['sha256'],'registros':len(rows),'quadros_graficos_preservados':unchanged,'layout':layout,'casos_com_sintese_revisada':len(analysis['casos']),'dados_gerados_por_ia':False})
    currenthash=hashlib.sha256((root/cfg['arquivo']).read_bytes()).hexdigest()
    assert currenthash==meta['sha256'],'Fonte alterada durante a geração.'
    old['analise_aprofundada']={'versao':'0.3.0','fonte_sha256':meta['sha256'],'redacao':'Textos técnicos e normativos; critérios operacionais no caderno de conferência.','referencia_federal':FEDERAL_SOURCE}
    old['arquivos']=[{'caminho':str(p.relative_to(run)),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in sorted(run.rglob('*')) if p.is_file() and p.name!='manifesto.json' and not p.name.startswith('~$')]
    json_write(run/'manifesto.json',old)
    (run.parent/'ULTIMA_EXECUCAO.txt').write_text(run.name,encoding='utf-8')
    print(json.dumps({'pasta':str(run),'redacao':layout,'casos':len(analysis['casos'])},ensure_ascii=False,indent=2))
    return run

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--config',default='config.local.json');ap.add_argument('--mesma-execucao',action='store_true');a=ap.parse_args()
    root=Path(__file__).resolve().parents[1];cfg=json.loads((root/a.config).read_text(encoding='utf-8-sig'));out=root/cfg['saida'];base=out/(out/'ULTIMA_EXECUCAO.txt').read_text().strip()
    generate(root,base,cfg,fork=not a.mesma_execucao)

if __name__=='__main__':main()
