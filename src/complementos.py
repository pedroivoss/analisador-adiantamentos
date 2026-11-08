"""Quadros mensais do consolidado e relação literal dos registros da fonte."""
from pathlib import Path
from collections import Counter
from copy import deepcopy
from datetime import date
from decimal import Decimal
import argparse,hashlib,json,subprocess,os,math
from analisador import analyze,load_workbook,cents,json_write
from apresentacao import moeda,mes

MONTH_NAMES=['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro']
DETAIL_COLS=['B','C','D','E','F','G','H','S','T','U','W','Z','AA']
DETAIL_HEADERS=['Servidor','Secretaria','Nº nota','Adiantado (R$)','Utilizado (R$)','Devolvido (R$)','Mês','Aplicação em atraso','Devolução em atraso','Prestação em atraso','Tipo','Exame SMF','Exame CGM']
def money(v):return 'R$ '+moeda(v)
def rate(a,b):return f'{a/b*100:.2f}%'.replace('.',',') if b else 'não apurável'
def listing(xs):return ', '.join(xs) if xs else 'nenhuma'

def guide(run,payload):
    from docx import Document
    from docx.shared import Inches,Pt,RGBColor
    from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer
    from reportlab.lib.styles import ParagraphStyle
    from xml.sax.saxutils import escape
    entries=[
        ('Relatório completo para o RCI','O Word editável e o PDF reúnem a análise semestral atualizada, as tabelas e gráficos, os seis quadros mensais e a presença das secretarias. Use o Word para selecionar e copiar os trechos ao RCI.'),
        ('Análise mensal e detalhada.xlsx','A planilha em Planilhas contém oito abas: janeiro a junho, presença mensal e relação individual. Os quadros mensais incluem as 21 secretarias, inclusive as sem movimento. Os totais são fórmulas e os valores permanecem numéricos.'),
        ('Relação individual dos adiantamentos','O Word e o PDF apresentam os 139 registros com os treze campos solicitados, comentários e quadros em paisagem. Os textos de mês, natureza, atraso e exame são reproduzidos como estão no controle.'),
        ('Mensal por secretaria','Cada pasta mensal contém o quadro em PNG e o texto que o acompanha. A pasta de presença mostra quais unidades receberam adiantamentos em todos, em parte ou em nenhum dos meses. As imagens revisadas permitem inserção manual no RCI.'),
        ('Relação individual e Textos analíticos','A pasta da relação individual contém as imagens por parte e os comentários. Textos analíticos reúne os comentários editáveis por assunto. O texto integral também fica em TXT para copiar sem formatação.'),
        ('Caderno de conferência e dados','O caderno mantém os critérios operacionais, referências normativas e pontos que exigem confirmação. Os arquivos de dados registram origem, conciliação, validações e hashes. Esse material serve à conferência e fica separado do corpo do RCI.'),
        ('Atualização e fidelidade','A execução usa a planilha definitiva recebida em 07/11/2026. Foram conferidos 1.807 campos literais, 378 valores mensais e 18 totais. A fonte permanece intacta. Correções posteriores devem ser feitas na fonte e seguidas de nova execução; editar uma saída não atualiza as demais.'),
    ]
    doc=Document();sec=doc.sections[0];sec.page_width=Inches(8.5);sec.page_height=Inches(11);sec.left_margin=sec.right_margin=Inches(66/72);sec.top_margin=sec.bottom_margin=Inches(.65)
    for name in ['Normal','Heading 1','Heading 2']:doc.styles[name].font.name='Arial';doc.styles[name].font.color.rgb=RGBColor(0,0,0)
    doc.styles['Normal'].font.size=Pt(10);doc.styles['Normal'].paragraph_format.space_after=Pt(7);doc.styles['Heading 1'].font.size=Pt(16);doc.styles['Heading 2'].font.size=Pt(11);doc.styles['Heading 2'].paragraph_format.space_before=Pt(8);doc.styles['Heading 2'].paragraph_format.space_after=Pt(3)
    head=ParagraphStyle('h',fontName='Helvetica-Bold',fontSize=15,leading=18,spaceAfter=12);sub=ParagraphStyle('s',fontName='Helvetica-Bold',fontSize=10.5,leading=13,spaceBefore=7,spaceAfter=3);body=ParagraphStyle('b',fontName='Helvetica',fontSize=9.8,leading=12.3,spaceAfter=6)
    title='Guia dos relatórios de adiantamentos';doc.add_paragraph(title,'Heading 1');story=[Paragraph(title,head)]
    for title,text in entries:
        doc.add_paragraph(title,'Heading 2');doc.add_paragraph(text);story.extend([Paragraph(escape(title),sub),Paragraph(escape(text),body)])
    doc.save(run/'Guia dos relatorios.docx');SimpleDocTemplate(str(run/'Guia dos relatorios.pdf'),pagesize=(612,792),leftMargin=66,rightMargin=66,topMargin=40,bottomMargin=40).build(story)

def prepare(root,run,cfg):
    rows,tables,s,meta=analyze(cfg,root)
    source=root/cfg['arquivo'];wb=load_workbook(source,data_only=True);ws=wb[cfg['aba']];cs=wb[cfg['aba_consolidado']]
    secs=[]
    for i in range(3,cs.max_row+1):
        sig=cs.cell(i,1).value
        if sig is not None and str(sig).strip() and str(sig).strip().upper()!='TOTAL':
            secs.append({'sigla':str(sig),'nome':str(cs.cell(i,2).value or ''),'linha':i})
    assert len({x['sigla'].strip().upper() for x in secs})==len(secs),'Siglas repetidas no consolidado.'
    months=[];checks=[]
    for month in tables['mensal']:
        mi=int(month['mes'][-2:]);base=3+(mi-1)*3;matrix=[]
        for sec in secs:
            vals=[]
            for j,key in enumerate(['adiantado','utilizado','devolvido']):
                cell=cs.cell(sec['linha'],base+j);actual=cents(cell.value)
                assert actual is not None,f'Consolidado sem valor armazenado: {cell.coordinate}'
                expected=sum(r[key] for r in rows if r['mes']==month['mes'] and r['secretaria']==sec['sigla'].strip().upper())
                assert actual==expected,f'Divergência entre abas em {cell.coordinate}: {actual} vs {expected}'
                vals.append(actual/100);checks.append({'celula':cell.coordinate,'centavos':actual,'conciliado':True})
            matrix.append([sec['sigla'],sec['nome'],*vals])
        totals=[sum(cents(r[j]) for r in matrix)/100 for j in [2,3,4]]
        for j,v in enumerate(totals):assert cents(v)==cents(cs.cell(24,base+j).value),'Total mensal diverge da linha 24.'
        active=[r for r in matrix if r[2]>0];inactive=[r[0] for r in matrix if r[2]==0]
        ordered=sorted(active,key=lambda x:-x[2]);top=ordered[0];largest_ret=max(active,key=lambda x:x[4]);top3=ordered[:3]
        prev=months[-1] if months else None
        comments=[
            f'Em {MONTH_NAMES[mi-1]} de {month["mes"][:4]}, {len(active)} das {len(secs)} secretarias apresentaram adiantamentos, totalizando {money(totals[0])}. A utilização foi de {money(totals[1])}, correspondente a {rate(totals[1],totals[0])} do valor concedido, e a devolução de {money(totals[2])}, equivalente a {rate(totals[2],totals[0])}.',
            f'A {top[0]} recebeu o maior montante do mês, {money(top[2])}, representando {rate(top[2],totals[0])} das concessões. As três maiores posições ({listing([r[0] for r in top3])}) reuniram {money(sum(r[2] for r in top3))}, ou {rate(sum(r[2] for r in top3),totals[0])}. A maior devolução em valor absoluto foi da {largest_ret[0]}, de {money(largest_ret[4])}.',
            f'As secretarias sem adiantamento registrado no mês foram: {listing(inactive)}. A ausência de movimentação delimita a distribuição dos recursos no período; não demonstra, por si só, inexistência de solicitação ou de necessidade administrativa.'
        ]
        if prev:
            delta=totals[0]-prev['totais'][0];sign='aumento' if delta>=0 else 'redução'
            prev_active={r[0] for r in prev['linhas'] if r[2]>0};now={r[0] for r in active}
            comments.append(f'Em relação a {prev["rotulo"]}, houve {sign} de {money(abs(delta))} ({rate(abs(delta),prev["totais"][0])}) no total adiantado. Passaram a apresentar movimento: {listing(sorted(now-prev_active))}. Deixaram de apresentar movimento: {listing(sorted(prev_active-now))}. O acompanhamento deve relacionar essas alterações às justificativas e aos eventos de cada unidade.')
        else:comments.append('A distribuição mensal permite confrontar a concentração das concessões com os saldos restituídos. O acompanhamento das estimativas deve considerar a finalidade e a excepcionalidade de cada pedido, sem interpretar o volume concedido isoladamente como evidência de irregularidade.')
        months.append({'chave':month['mes'],'rotulo':mes(month['mes']),'titulo':f'Adiantamentos por secretaria — {MONTH_NAMES[mi-1]}/{month["mes"][:4]}','linhas':matrix,'totais':totals,'comentarios':comments})
    presence=[]
    for sec in secs:
        vals=[next(r for r in m['linhas'] if r[0]==sec['sigla'])[2] for m in months]
        marks=['Com' if v>0 else 'Sem' for v in vals]
        presence.append([sec['sigla'],*marks,sum(v>0 for v in vals)])
    allmonths=[r[0] for r in presence if r[-1]==len(months)];none=[r[0] for r in presence if r[-1]==0]
    counts=Counter(r[-1] for r in presence)
    frequent=sum(cents(r[2]) for m in months for r in m['linhas'] if r[0] in allmonths)/100
    active_counts=[sum(r[2]>0 for r in m['linhas']) for m in months]
    most=[m['rotulo'] for m,c in zip(months,active_counts) if c==max(active_counts)]
    least=[m['rotulo'] for m,c in zip(months,active_counts) if c==min(active_counts)]
    presence_comments=[
        (f'A {allmonths[0]} foi a única secretaria com adiantamento em todos os {len(months)} meses. Recebeu ' if len(allmonths)==1 else f'As secretarias com adiantamento em todos os {len(months)} meses foram: {listing(allmonths)}. Juntas, receberam ')+f'{money(frequent)}, equivalentes a {rate(frequent,s["adiantado"])} dos recursos do período. A continuidade de registros recomenda acompanhar a recorrência das necessidades e a adequação do planejamento, preservada a análise individual das justificativas.',
        f'Não houve adiantamento registrado em nenhum mês para: {listing(none)}. Das {len(secs)} secretarias relacionadas no consolidado, {len(secs)-len(none)} tiveram movimento em pelo menos um mês. '+ 'A distribuição de presença foi: '+ '; '.join(f'{count} secretaria(s) em {n} mês(es)' for n,count in sorted(counts.items()))+'.',
        f'O maior alcance de secretarias ocorreu em {listing(most)}, com {max(active_counts)} unidades, e o menor em {listing(least)}, com {min(active_counts)}. A amplitude institucional deve ser lida junto ao volume financeiro: mais unidades participantes não significam, necessariamente, maior valor concedido.',
        '“Com” indica valor adiantado maior que zero no mês; “Sem” indica valor zero. Os meses sem movimento não autorizam concluir que pedidos tenham sido recusados ou que não tenham sido apresentados. O quadro trata dos adiantamentos realizados.'
    ]
    detail=[];provenance=[]
    for r in rows:
        i=r['linha_origem'];values=[ws[f'{c}{i}'].value for c in DETAIL_COLS]
        assert all(not isinstance(v,str) or not v.startswith('=') for v in values),'Fórmula sem resultado no detalhe'
        detail.append(values);provenance.append({'linha_fonte':i,'processo':r['processo'],'nota':r['nota']})
    # Os textos usam as contagens estruturadas do motor, não alteram os campos literais.
    raw_flags=[sum(str(v[j] or '').strip().lower().startswith('sim') for v in detail) for j in [7,8,9]]
    some=sum(any(str(v[j] or '').strip().lower().startswith('sim') for j in [7,8,9]) for v in detail)
    cgm=Counter(r['cgm'] for r in rows);smf=Counter(r['smf'] for r in rows)
    detail_comments=[
        f'A relação individual reúne {len(rows)} registros vinculados a {s["processos_unicos"]} processos SEI, com {money(s["adiantado"])} adiantados, {money(s["utilizado"])} utilizados e {money(s["devolvido"])} devolvidos. As {s["servidores"]} identificações nominais presentes no controle distribuem-se por {s["secretarias"]} secretarias. O fechamento financeiro é compatível com os totais mensais e com o consolidado do semestre.',
        f'Os campos de acompanhamento registram atraso de aplicação em {raw_flags[0]} contas, de devolução em {raw_flags[1]} e de prestação em {raw_flags[2]}. Considerando conjuntamente os três eventos, {some} registros apresentam ao menos uma marcação de atraso, equivalentes a {rate(some,len(rows))} da relação. Uma mesma conta pode aparecer em mais de um evento, razão pela qual as três contagens não devem ser somadas como pessoas ou processos distintos.',
        f'No exame da CGM, há {cgm["Aprovado"]} registros aprovados, {cgm["Ressalva"]} com ressalva e {cgm["Rejeitado"]} rejeitados. Na SMF, são {smf["Aprovado"]} aprovados e {smf["Ressalva"]} com ressalva; {smf["Ausente"]} campo(s) indicam ausência de exame e {smf["Não informado"]} não apresentam classificação. As {s["aprovados_ambos"]} aprovações por ambos os órgãos devem ser distinguidas das demais combinações, que exigem acompanhamento dos apontamentos e das providências adotadas.',
        'A leitura conjunta de responsável, secretaria, nota, valores, prazos, natureza e exames permite localizar as contas que concentram recursos e as que exigem acompanhamento. Repetição de servidor ou processo não caracteriza duplicidade automaticamente: um processo pode conter notas e naturezas distintas. Os resultados de ressalva e rejeição são preservados e não convertem o valor integral do adiantamento em dano ou débito.',
        'Os campos de atraso reproduzem as marcações do controle. A confirmação de intempestividade normativa deve observar os vencimentos aplicáveis, o calendário municipal e o marco de protocolo da prestação. A regularização documental e os eventuais ressarcimentos devem ser acompanhados por conta, conforme os arts. 13, 17 e 20 a 25 da IN-001/2018-CGM.'
    ]
    payload={'fonte':str(source),'sha256':meta['sha256'],'aba_consolidado':cfg['aba_consolidado'],'aba_detalhe':cfg['aba'],'meses':months,'presenca':presence,'comentarios_presenca':presence_comments,'campos_detalhe':DETAIL_HEADERS,'colunas_fonte':DETAIL_COLS,'detalhe':detail,'rastreio':provenance,'comentarios_detalhe':detail_comments,'conferencias':checks,'indicadores':s}
    json_write(run/'dados/complementos.json',payload)
    return payload

def documents(run,payload):
    from docx import Document
    from docx.shared import Inches,Pt,RGBColor
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from reportlab.platypus import SimpleDocTemplate,Paragraph,Table,TableStyle,PageBreak,Spacer
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from xml.sax.saxutils import escape
    def setup(landscape=False):
        d=Document();sec=d.sections[0];sec.page_width=Inches(11 if landscape else 8.5);sec.page_height=Inches(8.5 if landscape else 11)
        sec.left_margin=sec.right_margin=Inches(.5 if landscape else 66/72);sec.top_margin=sec.bottom_margin=Inches(.5 if landscape else .6)
        for name in ['Normal','Title','Heading 1','Heading 2']:
            d.styles[name].font.name='Arial';d.styles[name].font.color.rgb=RGBColor(0,0,0)
        d.styles['Normal'].font.size=Pt(10);d.styles['Normal'].paragraph_format.space_after=Pt(6);d.styles['Normal'].paragraph_format.line_spacing=1.05
        d.styles['Heading 1'].font.size=Pt(13);d.styles['Heading 1'].paragraph_format.space_after=Pt(8)
        d.core_properties.author='';return d
    ps=ParagraphStyle('Body',fontName='Helvetica',fontSize=9.3,leading=11.7,spaceAfter=6)
    hs=ParagraphStyle('Head',fontName='Helvetica-Bold',fontSize=13,leading=16,spaceAfter=9)
    def pp(t,style=ps):return Paragraph(escape(str(t)),style)
    def table(d,headers,records,widths,size=8.3,whole=True,moneycols=()):
        t=d.add_table(rows=1,cols=len(headers));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
        for i,w in enumerate(widths):t.columns[i].width=Inches(w/72)
        matrices=[headers,*records]
        for ri,row in enumerate(matrices):
            cells=t.rows[0].cells if ri==0 else t.add_row().cells
            trpr=t.rows[ri]._tr.get_or_add_trPr();trpr.append(OxmlElement('w:cantSplit'))
            if ri==0:trpr.append(OxmlElement('w:tblHeader'))
            for j,val in enumerate(row):
                c=cells[j];c.width=Inches(widths[j]/72);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
                c.text='' if val is None else str(val)
                tcpr=c._tc.get_or_add_tcPr();sh=OxmlElement('w:shd');sh.set(qn('w:fill'),'E3E3E3' if ri==0 else ('F5F5F5' if ri%2==0 else 'FFFFFF'));tcpr.append(sh)
                bd=OxmlElement('w:tcBorders')
                for edge in ['top','left','bottom','right']:
                    el=OxmlElement('w:'+edge);el.set(qn('w:val'),'single');el.set(qn('w:sz'),'4');el.set(qn('w:color'),'D9D9D9');bd.append(el)
                tcpr.append(bd)
                mar=OxmlElement('w:tcMar')
                for key,v in [('top',42),('bottom',42),('left',45),('right',45)]:
                    e=OxmlElement('w:'+key);e.set(qn('w:w'),str(v));e.set(qn('w:type'),'dxa');mar.append(e)
                tcpr.append(mar)
                for p in c.paragraphs:
                    p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=1.05
                    p.paragraph_format.keep_with_next=whole and ri<len(matrices)-1
                    p.alignment=WD_ALIGN_PARAGRAPH.CENTER if ri==0 or (j>0 and j not in moneycols) else (WD_ALIGN_PARAGRAPH.RIGHT if j in moneycols else WD_ALIGN_PARAGRAPH.LEFT)
                    if len(headers)==5 and j==1 and ri>0:p.alignment=WD_ALIGN_PARAGRAPH.LEFT
                    for r in p.runs:r.font.size=Pt(7 if ri==0 and len(headers)==13 else size);r.font.bold=ri==0
        return t
    def ptable(headers,records,widths,size=8):
        st=ParagraphStyle('Cell'+str(size),fontName='Helvetica',fontSize=size,leading=size*1.18)
        grid=[[pp(v,st) for v in headers]]+[[pp('' if v is None else v,st) for v in row] for row in records]
        t=Table(grid,colWidths=widths,repeatRows=1);t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E3E3E3')),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#F5F5F5')]),('GRID',(0,0),(-1,-1),.35,colors.HexColor('#D9D9D9')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3)]));return t
    md=setup();story=[];text=[]
    for ix,m in enumerate(payload['meses']):
        if ix:md.add_page_break();story.append(PageBreak())
        md.add_paragraph(m['titulo'],'Heading 1');story.append(pp(m['titulo'],hs));text.append(m['titulo'])
        headers=['Sigla','Secretaria','Adiantado (R$)','Utilizado (R$)','Devolvido (R$)'];widths=[47,187,82,82,82]
        vals=[[r[0],r[1],*[moeda(v) for v in r[2:]]] for r in m['linhas']]+[['Total','',*[moeda(v) for v in m['totais']]]]
        table(md,headers,vals,widths,size=8.0,moneycols=[2,3,4]);story.extend([ptable(headers,vals,widths,7.9),Spacer(1,8)])
        for p in m['comentarios']:md.add_paragraph(p);story.append(pp(p));text.append(p)
    md.add_page_break();story.append(PageBreak());title='Presença mensal das secretarias nos adiantamentos';md.add_paragraph(title,'Heading 1');story.append(pp(title,hs));text.append(title)
    headers=['Secretaria',*[m['rotulo'] for m in payload['meses']],'Meses com adiantamento'];widths=[67,*([53]*len(payload['meses'])),95]
    table(md,headers,payload['presenca'],widths,size=8.3);story.extend([ptable(headers,payload['presenca'],widths),Spacer(1,8)])
    for p in payload['comentarios_presenca']:md.add_paragraph(p);story.append(pp(p));text.append(p)
    md.save(run/'Analise mensal por secretaria.docx')
    SimpleDocTemplate(str(run/'Analise mensal por secretaria.pdf'),pagesize=(612,792),leftMargin=66,rightMargin=66,topMargin=38,bottomMargin=38).build(story)
    # A seção mensal também integra o corpo completo destinado ao RCI.
    main=Document(run/'Relatorio completo para o RCI.docx')
    # Evita duplicação quando o complemento é executado novamente na mesma pasta.
    children=list(main._element.body)
    for pos,element in enumerate(children):
        if ''.join(element.itertext()).startswith(payload['meses'][0]['titulo']):
            start=pos-1 if pos and children[pos-1].find('.//'+qn('w:br')) is not None else pos
            for old in children[start:]:
                if old.tag!=qn('w:sectPr'):main._element.body.remove(old)
            break
    main.add_page_break()
    for element in md._element.body:
        if element.tag!=qn('w:sectPr'):main._element.body.insert(len(main._element.body)-1,deepcopy(element))
    main.save(run/'Relatorio completo para o RCI.docx')
    with (run/'Texto integral para o RCI.txt').open('a',encoding='utf-8') as f:f.write('\n\n'+'\n\n'.join(text))
    # Relação larga: orientação paisagem e paginação por blocos inteiros.
    dd=setup(True);ds=[];title='Relação individual dos adiantamentos — janeiro a junho de 2026'
    dd.add_paragraph(title,'Heading 1');ds.append(pp(title,hs))
    for p in payload['comentarios_detalhe']:dd.add_paragraph(p);ds.append(pp(p))
    headers=['Servidor','Secre-\ntaria','Nº\nnota','Adiantado\n(R$)','Utilizado\n(R$)','Devolvido\n(R$)','Mês','Aplicação\nem atraso','Devolução\nem atraso','Prestação\nem atraso','Tipo','Exame\nSMF','Exame\nCGM']
    widths=[113,41,26,49,49,49,25,42,42,42,50,44,48]
    raw=payload['detalhe'];display=[]
    for row in raw:
        display.append([moeda(v) if j in [3,4,5] else ('' if v is None else str(v)) for j,v in enumerate(row)])
    batches=[];batch=[]
    for r in display:
        trial=batch+[r];pt=ptable(headers,trial,widths,7.5);height=pt.wrap(720,1000)[1]
        if height>418 and batch:batches.append(batch);batch=[r]
        else:batch=trial
    if batch:batches.append(batch)
    for i,records in enumerate(batches):
        dd.add_page_break();ds.append(PageBreak());label=f'Relação individual — parte {i+1} de {len(batches)}'
        dd.add_paragraph(label,'Heading 1');ds.append(pp(label,hs))
        table(dd,headers,records,widths,size=7.8,moneycols=[3,4,5]);ds.append(ptable(headers,records,widths,7.5))
        first=sum(len(b) for b in batches[:i])+1;last=first+len(records)-1;caption=f'Registros {first} a {last} de {len(raw)}.'
        dd.add_paragraph(caption);ds.extend([Spacer(1,5),pp(caption)])
    dd.save(run/'Relacao individual dos adiantamentos.docx')
    SimpleDocTemplate(str(run/'Relacao individual dos adiantamentos.pdf'),pagesize=(792,612),leftMargin=36,rightMargin=36,topMargin=36,bottomMargin=36).build(ds)
    (run/'Textos analíticos/Analise mensal por secretaria.txt').write_text('\n\n'.join(text),encoding='utf-8')
    (run/'Textos analíticos/Relacao individual.txt').write_text('\n\n'.join(payload['comentarios_detalhe']),encoding='utf-8')
    json_write(run/'dados/paginacao_detalhe.json',{'blocos':[len(b) for b in batches],'total':len(raw)})
    from pypdf import PdfWriter
    merged=PdfWriter()
    for p in [run/'Relatorio completo para o RCI.pdf',run/'Analise mensal por secretaria.pdf']:merged.append(p)
    merged.write(run/'Relatorio completo para o RCI.tmp.pdf');os.replace(run/'Relatorio completo para o RCI.tmp.pdf',run/'Relatorio completo para o RCI.pdf')

def main():
    root=Path(__file__).resolve().parents[1];ap=argparse.ArgumentParser();ap.add_argument('--config',default='config.local.json');a=ap.parse_args();cfg=json.loads((root/a.config).read_text(encoding='utf-8-sig'));out=root/cfg['saida'];run=out/(out/'ULTIMA_EXECUCAO.txt').read_text().strip()
    payload=prepare(root,run,cfg);documents(run,payload)
    node=os.environ.get('ADIANTAMENTO_NODE','node');subprocess.run([node,str(root/'src/exportar_complementos.mjs'),str(run)],check=True)
    from imprimir_planilha import configure_print
    configure_print(run/'Planilhas/Analise mensal e detalhada.xlsx',len(payload['detalhe']))
    guide(run,payload)
    assert hashlib.sha256((root/cfg['arquivo']).read_bytes()).hexdigest()==payload['sha256'],'Fonte alterada durante geração'
    manifest=json.loads((run/'manifesto.json').read_text(encoding='utf-8'))
    manifest['complementos']={'meses':len(payload['meses']),'secretarias_no_consolidado':len(payload['presenca']),'campos_literais_por_registro':13,'fonte_sha256':payload['sha256']}
    manifest['arquivos']=[{'caminho':str(p.relative_to(run)),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in sorted(run.rglob('*')) if p.is_file() and p.name!='manifesto.json' and not p.name.startswith('~$')]
    json_write(run/'manifesto.json',manifest)
    print(json.dumps({'pasta':str(run),'meses':len(payload['meses']),'secretarias_consolidado':len(payload['presenca']),'registros':len(payload['detalhe']),'conferencias':len(payload['conferencias'])},ensure_ascii=False))
if __name__=='__main__':main()
