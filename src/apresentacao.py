"""Quadros para o RCI. Uma definição alimenta Excel, Word, PDF e imagens."""
from pathlib import Path
from collections import Counter,defaultdict
from datetime import date,datetime
from decimal import Decimal
import json,re,hashlib
from openpyxl import load_workbook

WIDTH=480
MONTHS=['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez']
def mes(value):
    y,m=value[:7].split('-');return f'{MONTHS[int(m)-1]}/{y[2:]}'
def moeda(value):
    if value is None:return 'Não informado'
    return f'{value:,.2f}'.replace(',','@').replace('.',',').replace('@','.')
def display(value,fmt='texto'):
    if value is None:return 'Não disponível'
    if fmt=='moeda':return moeda(value)
    if fmt=='percentual':return f'{value*100:.2f}%'.replace('.',',')
    if fmt=='decimal':return f'{value:.1f}'.replace('.',',')
    return str(value)

def cell_display(plan,i,j):
    fmt='moeda' if plan['id']=='Resumo' and i>=4 and j==1 else plan['formatos'][j]
    return display(plan['linhas'][i][j],fmt)

def prepare(root,run,rows,tables,summary,meta):
    from analisador import cents,norm,status,stats,json_write
    source=root/meta['config']['arquivo'];ws=load_workbook(source,data_only=True)[meta['aba']]
    raw=[]
    cols={'A':'Processo SEI','B':'Servidor','C':'Secretaria','D':'Nº da nota','E':'Valor adiantado','F':'Valor utilizado','G':'Valor devolvido','H':'Mês informado','I':'Data do pagamento','K':'Prazo de aplicação','L':'Prazo de devolução','M':'Prazo da prestação','O':'Última aplicação','P':'Data da devolução','Q':'Data da prestação','S':'Aplicação em atraso','T':'Devolução em atraso','U':'Prestação em atraso','W':'Natureza','Z':'Exame SMF','AA':'Exame CGM'}
    for r in rows:
        raw.append({'Linha na fonte':r['linha_origem'],**{label:ws[f'{col}{r["linha_origem"]}'].value for col,label in cols.items()}})
    def date_json(x):return x.strftime('%d/%m/%Y') if isinstance(x,(date,datetime)) else x
    json_write(run/'dados/transcricao_fonte.json',[{k:date_json(v) for k,v in r.items()} for r in raw])
    checks=[]
    for key,label in [('adiantado','Valor adiantado'),('utilizado','Valor utilizado'),('devolvido','Valor devolvido')]:
        amount=sum(cents(r[label]) for r in raw if cents(r[label]) is not None)
        assert amount==cents(summary[key]),f'Total divergente: {label}'
        checks.append({'conferencia':label,'centavos_fonte':amount,'centavos_relatorio':cents(summary[key]),'confere':True})
    assert len(raw)==summary['registros']
    for group_key,original_label,table_key in [('secretaria','Secretaria','secretarias'),('servidor','Servidor','servidores')]:
        for item in tables[table_key]:
            matching=[r for r in raw if ' '.join(str(r[original_label] or '').upper().split())==item[group_key]]
            assert len(matching)==item['registros'],f'Quantidade divergente: {item[group_key]}'
            for key,label in [('adiantado','Valor adiantado'),('utilizado','Valor utilizado'),('devolvido','Valor devolvido')]:
                assert sum(cents(r[label]) for r in matching if cents(r[label]) is not None)==cents(item[key]),f'Grupo divergente: {item[group_key]} / {label}'
    names={r['linha_origem']:str(ws[f'B{r["linha_origem"]}'].value or '') for r in rows}
    name_map={r['servidor']:names[r['linha_origem']] for r in rows}
    plans=[]
    def add(key,title,headers,data,widths,formats,why,conclusion,method,chart=None,total=False):
        assert len(headers)<=6 and sum(widths)==WIDTH
        assert len(headers)==len(formats)==len(widths)
        assert all(len(row)==len(headers) for row in data)
        assert not any('_' in h for h in headers)
        plans.append({'id':key,'titulo':title,'cabecalhos':headers,'linhas':data,'larguras':widths,'formatos':formats,'o_que_mostra':why,'conclusao':conclusion,'criterio':method,'grafico':chart,'tem_total':total})
    s=summary;total=s['adiantado'];ratio=lambda n,d:n/d if d else None
    add('Resumo','Panorama dos adiantamentos', ['Indicador','Resultado'],[
        ['Registros por nota e natureza',s['registros']],['Processos SEI distintos',s['processos_unicos']],['Servidores na base',s['servidores']],['Secretarias na base',s['secretarias']],
        ['Valor adiantado (R$)',s['adiantado']],['Valor utilizado (R$)',s['utilizado']],['Valor devolvido (R$)',s['devolvido']],['Diferença financeira (R$)',s['diferenca_financeira']]],
        [330,150],['texto','texto'],'Dimensiona os registros disponíveis e apresenta a composição financeira do período.',
        f'Os valores utilizados ({moeda(s["utilizado"])} reais) e devolvidos ({moeda(s["devolvido"])} reais) somam o valor adiantado de {moeda(total)} reais. A diferença financeira agregada é de {moeda(s["diferenca_financeira"])} reais.',
        'Contagens e somas das linhas incluídas. Um processo pode conter mais de uma nota; registros não são a quantidade exata de solicitações.')
    money_headers=['Mês','Adiantado (R$)','Utilizado (R$)','Devolvido (R$)']
    monthly=[[mes(r['mes']),r['adiantado'],r['utilizado'],r['devolvido']] for r in tables['mensal']]
    monthly.append(['Total',s['adiantado'],s['utilizado'],s['devolvido']])
    high=max(tables['mensal'],key=lambda r:r['adiantado']);low=min(tables['mensal'],key=lambda r:r['adiantado'])
    add('Mensal','Evolução mensal dos valores',money_headers,monthly,[90,130,130,130],['texto','moeda','moeda','moeda'],
        'Compara os valores adiantados, utilizados e devolvidos, agrupados pelo mês do pagamento.',
        f'O maior valor adiantado foi de R$ {moeda(high["adiantado"])} em {mes(high["mes"])}; o menor foi de R$ {moeda(low["adiantado"])} em {mes(low["mes"])}.',
        'Soma das colunas valor adt, valor utl e valor dev. Mês derivado de DATA PAGTO. Datas posteriores ao semestre permanecem vinculadas ao pagamento de origem.',chart='mensal',total=True)
    secs=tables['secretarias'];top=secs[0]
    add('Secretarias','Valores e registros por secretaria',['Secretaria','Registros','Adiantado (R$)','Utilizado (R$)','Devolvido (R$)'],
        [[r['secretaria'],r['registros'],r['adiantado'],r['utilizado'],r['devolvido']] for r in secs]+[['Total',s['registros'],s['adiantado'],s['utilizado'],s['devolvido']]],
        [68,55,119,119,119],['texto','inteiro','moeda','moeda','moeda'],
        'Apresenta todas as secretarias com registros na base, ordenadas pelo valor adiantado.',
        f'{top["secretaria"]} concentra o maior valor adiantado, R$ {moeda(top["adiantado"])}, em {top["registros"]} registros. A concentração não demonstra irregularidade por si só.',
        'Agrupamento por SECRETARIA; somas e contagens de registros por nota. Siglas conforme a fonte, com apenas espaços e caixa padronizados na consolidação.',total=True)
    nature=[]
    for r in tables['naturezas']:
        label=r['tipo']
        if label.startswith('Outro'):label='Outra natureza'
        nature.append([label,r['registros'],r['adiantado'],r['utilizado'],r['devolvido']])
    other=sorted({str(r['Natureza']) for r in raw if norm(r['Natureza']) not in ['servico','servicos','consumo']})
    add('Naturezas','Distribuição por natureza',['Natureza','Registros','Adiantado (R$)','Utilizado (R$)','Devolvido (R$)'],nature,
        [86,55,113,113,113],['texto','inteiro','moeda','moeda','moeda'],
        'Segrega os valores conforme a natureza informada na planilha.',
        f'Há {s["sem_uso"]} registros com valor utilizado igual a zero e {s["devolucao_integral"]} com devolução integral aritmeticamente compatível. '+('A outra natureza informada é: '+ '; '.join(other)+'.' if other else ''),
        'Somente caixa e espaços são normalizados em serviço/consumo. A classificação diferente é preservada na transcrição e não é reclassificada.',chart='naturezas')
    servers=tables['servidores']
    def rank(rs,key):
        rs=sorted(rs,key=lambda r:(-r[key],r['servidor']));cut=rs[min(9,len(rs)-1)][key];return [r for r in rs if r[key]>=cut]
    most=rank(servers,'adiantado')
    add('Maiores valores','Servidores com maiores valores adiantados',['Servidor','Adiantado (R$)','Utilizado (R$)','Devolvido (R$)'],
        [[name_map[r['servidor']],r['adiantado'],r['utilizado'],r['devolvido']] for r in most],
        [255,75,75,75],['texto','moeda','moeda','moeda'],
        'Destaca os dez maiores totais adiantados por servidor, incluindo eventuais empates no corte.',
        f'O maior total adiantado pertence a {name_map[servers[0]["servidor"]]}, com R$ {moeda(servers[0]["adiantado"])}. Os valores representam todos os registros desse servidor no recorte.',
        'Somas por SERVIDOR. Nomes exibidos com a grafia da fonte. O levantamento não mede gastos pessoais, mas valores sob responsabilidade do servidor.')
    most=rank(servers,'registros');leader=most[0]['registros'];leaders=[name_map[r['servidor']] for r in most if r['registros']==leader]
    add('Mais registros','Servidores com maior número de registros',['Servidor','Registros','Processos SEI'],
        [[name_map[r['servidor']],r['registros'],r['processos_unicos']] for r in most],
        [330,75,75],['texto','inteiro','inteiro'],
        'Mostra a frequência de registros por servidor e a quantidade de processos SEI distintos correspondentes.',
        ', '.join(leaders)+f' apresentam a maior frequência, com {leader} registros por servidor. Um SEI com notas de naturezas diferentes conta uma vez na coluna de processos.',
        'Dez maiores quantidades, preservando empates no corte. Quantidade de registros não equivale necessariamente a pedidos distintos.')
    fin=s['adiantado_estatisticas']
    add('Valores por registro','Estatísticas do valor adiantado',['Indicador','Valor (R$)'],
        [['Média',fin['media']],['Mediana',fin['mediana']],['Menor valor',fin['minimo']],['Maior valor',fin['maximo']],['Percentil 90',fin['p90']]],
        [330,150],['texto','moeda'],
        'Resume o tamanho dos adiantamentos registrados, usando somente valores preenchidos.',
        f'A média é de R$ {moeda(fin["media"])}, enquanto a mediana é de R$ {moeda(fin["mediana"])}. O intervalo observado vai de R$ {moeda(fin["minimo"])} a R$ {moeda(fin["maximo"])}.',
        f'Indicadores calculados a partir de {fin["n"]} valores da coluna valor adt; não são campos originalmente preenchidos. Percentil 90 pela posição mais próxima.')
    events={'aplicacao':'Última aplicação','devolucao':'Devolução','prestacao':'Prestação'}
    durations=[[events[r['evento']],r['n'],r['media'],r['mediana'],r['minimo'],r['maximo']] for r in tables['tempos']]
    add('Durações','Tempo entre o pagamento e os eventos',['Evento','Datas válidas','Média (dias)','Mediana (dias)','Mínimo (dias)','Máximo (dias)'],durations,
        [120,64,74,74,74,74],['texto','inteiro','decimal','decimal','inteiro','inteiro'],
        'Apresenta a duração em dias corridos entre o pagamento e cada evento com datas válidas.',
        'As médias têm amostras diferentes, indicadas na tabela. A última aplicação não informa quando começou a utilização do dinheiro. Datas ausentes, anteriores ao pagamento ou futuras à observação não entram nessas médias.',
        'Indicadores derivados das datas da fonte. Devolução com saldo zero e aplicação sem uso são não aplicáveis. Os valores não foram produzidos por IA.',chart='prazos')
    delay=[];delay_detail=[]
    for col,label in [('Aplicação em atraso','Última aplicação'),('Devolução em atraso','Devolução'),('Prestação em atraso','Prestação')]:
        counts=Counter();days=[]
        for r in raw:
            v=norm(r[col]);match=re.search(r'(\d+)\s*dia',v)
            if v=='nao':counts['não']+=1
            elif v.startswith('sim'):
                counts['sim']+=1
                if match:
                    days.append(int(match.group(1)));delay_detail.append({'Evento':label,'Servidor':r['Servidor'],'Dias informados':int(match.group(1)),'Nº da nota':r['Nº da nota'],'Linha na fonte':r['Linha na fonte']})
            else:counts['outros']+=1
        delay.append([label,counts['não'],counts['sim'],counts['outros'],max(days) if days else None])
        assert sum(counts.values())==len(raw)
    add('Atrasos da fonte','Atrasos conforme informados na planilha',['Evento','Não','Sim','Outros','Maior atraso (dias)'],delay,
        [144,59,59,59,159],['texto','inteiro','inteiro','inteiro','inteiro'],
        'Reproduz as classificações Não, Sim e demais conteúdos das três colunas de atraso da fonte.',
        '; '.join(f'{r[0]}: {r[2]} registros marcados como Sim' for r in delay)+'. Esses resultados descrevem os campos preenchidos, sem substituir os vencimentos adotados pela responsável.',
        'Contagem direta de S, T e U. Outros inclui sem uso, campos ausentes e textos diferentes. O maior atraso é extraído do número de dias escrito em Sim. Não se aplica um calendário novo para modificar a fonte.')
    maxima=[]
    for event in events.values():
        ds=[r for r in delay_detail if r['Evento']==event]
        if ds:
            maximum=max(r['Dias informados'] for r in ds)
            maxima.extend([[r['Evento'],r['Servidor'],r['Dias informados'],r['Nº da nota']] for r in ds if r['Dias informados']==maximum])
    add('Maiores atrasos','Maiores atrasos informados por evento',['Evento','Servidor','Dias','Nº da nota'],maxima,
        [95,285,45,55],['texto','texto','inteiro','texto'],
        'Identifica os registros que contêm o maior número de dias de atraso escrito na planilha para cada evento.',
        'Os nomes e notas permitem localizar os registros. Os máximos não representam uma nova decisão sobre os processos e devem ser lidos com os vencimentos informados na fonte.',
        'Valores extraídos dos textos de S, T e U; todos os empates são preservados. Datas e pareceres originais não são alterados.')
    for organ in ['SMF','CGM']:
        rs=[r for r in tables['exames'] if r['orgao']==organ]
        add('Exame '+organ,'Resultados do exame '+organ,['Resultado','Registros','Participação','Adiantado (R$)'],
            [[r['resultado'],r['registros'],r['percentual'],r['adiantado']] for r in rs]+[['Total',len(raw),1,total]],
            [150,70,100,160],['texto','inteiro','percentual','moeda'],
            f'Consolida as categorias preenchidas no campo EXAME {organ} e os valores adiantados dos registros correspondentes.',
            '; '.join(f'{r["resultado"]}: {r["registros"]} registros' for r in rs)+'. Ressalva é apresentada separadamente de aprovação sem ressalva.',
            'Contagem por categoria, com espaços e caixa uniformizados. Percentuais calculados sobre todos os registros, incluindo ausências. Valores envolvidos não equivalem automaticamente a dano ou ressarcimento.',total=True)
    smfs=sorted({r['smf'] for r in rows});cgms=sorted({r['cgm'] for r in rows})
    cross=[[a,*[sum(r['smf']==a and r['cgm']==b for r in rows) for b in cgms],sum(r['smf']==a for r in rows)] for a in smfs]
    cross.append(['Total',*[sum(r['cgm']==b for r in rows) for b in cgms],len(rows)])
    add('Exames cruzados','Comparação entre os exames SMF e CGM',['Exame SMF',*['CGM: '+x for x in cgms],'Total'],cross,
        [140,*([85]*len(cgms)),480-140-85*len(cgms)],['texto']+['inteiro']*(len(cgms)+1),
        'Relaciona o resultado preenchido por cada órgão para os mesmos registros.',
        f'Há {s["aprovados_ambos"]} registros aprovados por ambos os órgãos. As categorias diferem em {s["exames_divergentes"]} registros, incluindo aqueles com exame ausente ou não informado.',
        'Linhas representam a SMF; colunas representam a CGM. Os exames são de prestação de contas, não a autorização inicial dos pedidos.',total=True)
    quality=Counter(x['codigo'] for x in tables['alertas'])
    labels={'PRAZO_DIVERGENTE_APLICACAO':'Prazo de aplicação diferente da referência técnica','PRAZO_DIVERGENTE_DEVOLUCAO':'Prazo de devolução diferente da referência técnica','PRAZO_DIVERGENTE_PRESTACAO':'Prazo da prestação diferente da referência técnica','EVENTO_ANTES_PAGAMENTO':'Data de evento anterior ao pagamento','DEVOLUCAO_CONTRADITORIA':'Valor e data de devolução contraditórios','DATA_EVENTO_AUSENTE':'Data de evento ausente ou inválida','EXAME_AUSENTE_OUTRO':'Exame ausente ou não padronizado','TIPO_NAO_PADRONIZADO':'Natureza fora de serviço e consumo','APLICACAO_SEM_VALOR':'Data de aplicação com utilização igual a zero'}
    add('Revisão dos dados','Pontos para revisão antes da versão definitiva',['Ponto de conferência','Ocorrências'],
        [[labels.get(k,k.replace('_',' ').capitalize()),v] for k,v in quality.most_common()],
        [390,90],['texto','inteiro'],
        'Organiza as verificações técnicas que precisam de conferência na fonte. Esta seção é de revisão e não constitui conclusão de irregularidade.',
        f'Foram emitidos {s["alertas"]} alertas em {s["linhas_com_alerta"]} linhas. As células preenchidas do consolidado têm {s["conciliacao_divergencias"]} diferenças em relação ao recálculo. Há {s["conciliacao_sem_cache"]} células sem valor armazenado.',
        'Uma linha pode ter vários alertas. Comparações de prazo são separadas dos atrasos escritos na planilha. O calendário municipal não está validado; não se emite julgamento normativo automático.')
    for start in range(0,len(servers),18):
        page=start//18+1;chunk=servers[start:start+18]
        add(f'Servidores {page}',f'Valores por servidor parte {page}',
            ['Servidor','Adiantado (R$)','Utilizado (R$)','Devolvido (R$)'],
            [[name_map[r['servidor']],r['adiantado'],r['utilizado'],r['devolvido']] for r in chunk],
            [255,75,75,75],['texto','moeda','moeda','moeda'],
            f'Relação completa dos servidores, posições {start+1} a {start+len(chunk)}, ordenada pelo total adiantado.',
            'Esta relação permite consultar os valores de todos os servidores presentes na base. As partes devem ser consideradas em conjunto para a leitura completa.',
            'Somas por servidor a partir das linhas incluídas. Valores em reais; grafia dos nomes preservada da fonte.')
    json_write(run/'dados/apresentacao.json',{'largura_pt':WIDTH,'quadros':plans,'fonte':meta['arquivo'],'aba':meta['aba'],'periodo':meta['config']})
    json_write(run/'dados/conferencia_apresentacao.json',{'conferencias':checks,'linhas_conferidas':len(raw),'sha256_fonte':hashlib.sha256(source.read_bytes()).hexdigest(),'nomes_preservados':True,'atrasos_principais':'Campos preenchidos na fonte S, T e U','numeros_gerados_por_ia':False})
    return plans

def create_report(root,run,plans,meta):
    import os,sys,subprocess,shutil
    from xml.sax.saxutils import escape
    from reportlab.pdfgen import canvas
    from reportlab.platypus import Table,TableStyle,Paragraph
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from docx import Document
    from docx.shared import Pt,Inches,RGBColor
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.text import WD_BREAK
    from docx.enum.table import WD_TABLE_ALIGNMENT,WD_CELL_VERTICAL_ALIGNMENT
    from analisador import json_write
    cfg=meta['config'];figures=run/'Quadros para o RCI';figures.mkdir(exist_ok=True)
    qa=root/'work'/run.name;qa.mkdir(parents=True,exist_ok=True)
    pdftoppm=Path(sys.executable).parents[1]/'native/poppler/Library/bin/pdftoppm.exe'
    if not pdftoppm.exists():pdftoppm=Path(shutil.which('pdftoppm') or '')
    if not pdftoppm.is_file():raise RuntimeError('Poppler não encontrado para criar as imagens das tabelas.')
    body=ParagraphStyle('Corpo',fontName='Helvetica',fontSize=10,leading=13,spaceAfter=6)
    small=ParagraphStyle('Nota',fontName='Helvetica',fontSize=8,leading=10)
    title_style=ParagraphStyle('Titulo',fontName='Helvetica-Bold',fontSize=15,leading=18)
    cell=ParagraphStyle('Celula',fontName='Helvetica',fontSize=8.4,leading=10.5)
    head=ParagraphStyle('Cabecalho',fontName='Helvetica-Bold',fontSize=8.4,leading=10.5,textColor=colors.white)
    def para(text,style=body):return Paragraph(escape(text),style)
    def make_table(plan):
        data=[[para(x,head) for x in plan['cabecalhos']]]
        for ri,row in enumerate(plan['linhas']):
            formatted=[]
            for j,v in enumerate(row):
                text=cell_display(plan,ri,j);sty=ParagraphStyle('C',parent=cell,alignment=0 if j==0 or plan['formatos'][j]=='texto' else 2)
                formatted.append(para(text,sty))
                if j>0 and plan['formatos'][j] in ['moeda','inteiro','decimal','percentual']:
                    assert stringWidth(text,'Helvetica',8.4)<=plan['larguras'][j]-10,f'Número não cabe: {plan["id"]} / {text}'
            data.append(formatted)
        table=Table(data,colWidths=plan['larguras'],repeatRows=1)
        rules=[('BACKGROUND',(0,0),(-1,0),colors.HexColor('#24476B')),('GRID',(0,0),(-1,-1),.35,colors.HexColor('#D9D9D9')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4)]
        for i in range(2,len(data),2):rules.append(('BACKGROUND',(0,i),(-1,i),colors.HexColor('#F2F5F8')))
        if plan['tem_total']:rules.extend([('BACKGROUND',(0,len(data)-1),(-1,len(data)-1),colors.HexColor('#E2EAF1'))])
        table.setStyle(TableStyle(rules));return table
    def new_doc():
        doc=Document();sec=doc.sections[0];sec.page_width=Inches(8.5);sec.page_height=Inches(11);sec.left_margin=sec.right_margin=Inches(66/72);sec.top_margin=sec.bottom_margin=Inches(.55)
        for name in ['Normal','Title','Heading 1','Heading 2']:
            st=doc.styles[name];st.font.name='Arial';st.font.color.rgb=RGBColor(0,0,0)
        doc.styles['Normal'].font.size=Pt(10);doc.styles['Normal'].paragraph_format.space_after=Pt(6);doc.styles['Normal'].paragraph_format.line_spacing=1.08
        doc.styles['Title'].font.size=Pt(15);doc.styles['Title'].font.bold=True;doc.styles['Title'].paragraph_format.space_after=Pt(6)
        for style in doc.styles:
            for border in style.element.findall('.//'+qn('w:pBdr')):border.getparent().remove(border)
        doc.core_properties.title='Adiantamentos janeiro a junho de 2026';doc.core_properties.author=''
        return doc
    editable=new_doc();images=new_doc()
    source=f'Fonte: {meta["arquivo"]}, aba {meta["aba"]}.'
    period=f'Pagamentos de {date.fromisoformat(cfg["inicio"]).strftime("%d/%m/%Y")} a {date.fromisoformat(cfg["fim"]).strftime("%d/%m/%Y")}. Dados observados em {date.fromisoformat(cfg["data_observacao"]).strftime("%d/%m/%Y")}. Preliminar.'
    pdf=canvas.Canvas(str(run/'Relatorio para o RCI.pdf'),pagesize=(612,792));pdf.setTitle('Adiantamentos para o RCI')
    heights=[]
    for index,plan in enumerate(plans,1):
        table=make_table(plan);_,th=table.wrap(WIDTH,700)
        # Imagem do quadro: título, tabela e fonte, sem textos técnicos.
        heading=para(plan['titulo'],title_style);_,hh=heading.wrap(WIDTH,500)
        foot=para(source,small);_,fh=foot.wrap(WIDTH,500)
        image_pdf=qa/(f'quadro-{index:02}.pdf');height=hh+th+fh+38
        c=canvas.Canvas(str(image_pdf),pagesize=(WIDTH+24,height));heading.drawOn(c,12,height-12-hh);table.drawOn(c,12,12+fh+8);foot.drawOn(c,12,10);c.save()
        prefix=figures/f'{index:02} {plan["id"]}'
        subprocess.run([str(pdftoppm),'-singlefile','-r','180','-png',str(image_pdf),str(prefix)],check=True,stdout=subprocess.DEVNULL,stderr=subprocess.PIPE)
        plan['imagem']=str(prefix.relative_to(run)).replace('\\','/')
        png=Path(str(prefix)+'.png')
        for doc in [editable,images]:
            if index>1:doc.add_page_break()
            doc.add_paragraph(plan['titulo'],'Title');doc.add_paragraph(period)
            doc.add_paragraph(plan['o_que_mostra'])
        # Tabela nativa com larguras fixas; linhas indivisíveis e cabeçalho repetível.
        t=editable.add_table(rows=1,cols=len(plan['cabecalhos']));t.autofit=False;t.alignment=WD_TABLE_ALIGNMENT.CENTER
        for col,width in zip(t.columns,plan['larguras']):col.width=Inches(width/72)
        for j,h in enumerate(plan['cabecalhos']):t.rows[0].cells[j].text=h
        repeat=OxmlElement('w:tblHeader');t.rows[0]._tr.get_or_add_trPr().append(repeat)
        for ri,row in enumerate(plan['linhas']):
            cells=t.add_row().cells
            for j,v in enumerate(row):cells[j].text=cell_display(plan,ri,j)
        for ri,row in enumerate(t.rows):
            ns=OxmlElement('w:cantSplit');row._tr.get_or_add_trPr().append(ns)
            for j,c1 in enumerate(row.cells):
                c1.width=Inches(plan['larguras'][j]/72);c1.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
                props=c1._tc.get_or_add_tcPr();sh=OxmlElement('w:shd');sh.set(qn('w:fill'),'24476B' if ri==0 else ('F2F5F8' if ri%2==0 else 'FFFFFF'));props.append(sh)
                borders=OxmlElement('w:tcBorders')
                for edge in ['top','left','bottom','right']:
                    el=OxmlElement('w:'+edge);el.set(qn('w:val'),'single');el.set(qn('w:sz'),'4');el.set(qn('w:color'),'D9D9D9');borders.append(el)
                props.append(borders)
                margins=OxmlElement('w:tcMar')
                for edge,val in [('top',55),('bottom',55),('left',70),('right',70)]:
                    el=OxmlElement('w:'+edge);el.set(qn('w:w'),str(val));el.set(qn('w:type'),'dxa');margins.append(el)
                props.append(margins)
                for p in c1.paragraphs:
                    p.paragraph_format.space_after=Pt(0);p.paragraph_format.space_before=Pt(0);p.paragraph_format.line_spacing=1;p.paragraph_format.keep_with_next=False
                    if j>0 and plan['formatos'][j]!='texto':p.alignment=2
                    for run1 in p.runs:run1.font.size=Pt(8.4);run1.font.bold=(ri==0);run1.font.color.rgb=RGBColor(255,255,255) if ri==0 else RGBColor(0,0,0)
        images.add_picture(str(png),width=Inches(WIDTH/72))
        for doc in [editable,images]:
            for label,text in [('Leitura dos resultados',plan['conclusao']),('Critério',plan['criterio'])]:
                p=doc.add_paragraph();p.paragraph_format.space_before=Pt(7 if label=='Leitura dos resultados' else 0);p.add_run(label+'. ').bold=True;p.add_run(text)
            p=doc.add_paragraph(source)
            for r in p.runs:r.font.size=Pt(8)
        # PDF com o mesmo conteúdo e um quadro por página.
        y=750
        def put(text,style=body,space=7):
            nonlocal y
            obj=para(text,style);_,h=obj.wrap(WIDTH,700);y-=h;obj.drawOn(pdf,66,y);y-=space
        put(plan['titulo'],title_style);put(period,small);put(plan['o_que_mostra'])
        table=make_table(plan);_,th=table.wrap(WIDTH,700)
        y-=th;table.drawOn(pdf,66,y);y-=10
        put('Leitura dos resultados. '+plan['conclusao']);put('Critério. '+plan['criterio'],small);put(source,small)
        if plan['grafico']:
            pic=run/'graficos'/(plan['grafico']+'.png');iw,ih=ImageReader(str(pic)).getSize();h=min(160,y-55);w=h*iw/ih
            if w>WIDTH:w=WIDTH;h=w*ih/iw
            assert h>=100,f'Gráfico sem espaço: {plan["id"]}'
            pdf.drawImage(str(pic),66+(WIDTH-w)/2,y-h,width=w,height=h);y-=h
            for doc in [editable,images]:doc.add_picture(str(pic),width=Inches(min(WIDTH,w)/72))
        assert y>=40,f'Conteúdo fora da página: {plan["id"]} ({y})'
        heights.append({'quadro':plan['id'],'tabela_pt':th,'margem_inferior_pt':y,'largura_pt':WIDTH})
        pdf.setFont('Helvetica',8);pdf.drawString(66,24,'Adiantamentos - análise preliminar');pdf.drawRightString(546,24,str(index));pdf.showPage()
    pdf.save();editable.save(run/'Relatorio para o RCI editavel.docx');images.save(run/'Relatorio para o RCI com imagens.docx')
    json_write(run/'dados/verificacao_layout.json',{'largura_tabelas_pt':WIDTH,'paginas_planejadas':len(plans),'quadros':heights})
    json_write(run/'dados/apresentacao.json',{'largura_pt':WIDTH,'quadros':plans,'fonte':meta['arquivo'],'aba':meta['aba'],'periodo':cfg})
