"""Análise reproduzível de prestações de contas. Nenhum cálculo depende de IA."""
from __future__ import annotations
import argparse, csv, hashlib, json, math, os, re, statistics, subprocess, sys, unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
_local_deps=Path(__file__).resolve().parents[1]/'.runtime-deps'
if _local_deps.exists():sys.path.insert(0,str(_local_deps))
from openpyxl import load_workbook

VERSION='0.1.0'
COLS={'A':'processo','B':'servidor','C':'secretaria','D':'nota','E':'adiantado','F':'utilizado','G':'devolvido','H':'mes_informado','I':'pagamento','K':'prazo_aplicacao','L':'prazo_devolucao','M':'prazo_prestacao','O':'ultima_aplicacao','P':'devolucao','Q':'prestacao','S':'atraso_informado_aplicacao','T':'atraso_informado_devolucao','U':'atraso_informado_prestacao','W':'tipo','Z':'smf','AA':'cgm'}
MONEY=['adiantado','utilizado','devolvido']
DATES=['pagamento','prazo_aplicacao','prazo_devolucao','prazo_prestacao','ultima_aplicacao','devolucao','prestacao']
MONTHS=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']

def norm(value):
    return ' '.join(''.join(c for c in unicodedata.normalize('NFKD',str(value or '')) if not unicodedata.combining(c)).lower().split())

def cents(value):
    if value is None or isinstance(value,bool): return None
    try:
        if isinstance(value,str):
            value=value.strip().replace('R$','').replace(' ','')
            if ',' in value: value=value.replace('.','').replace(',','.')
        d=Decimal(str(value))
        if not d.is_finite(): return None
        return int((d*100).quantize(Decimal('1'),rounding=ROUND_HALF_UP))
    except Exception: return None

def dt(value):
    if isinstance(value,datetime): return value.date()
    if isinstance(value,date): return value
    if isinstance(value,str):
        for fmt in ['%d/%m/%Y','%Y-%m-%d','%Y-%m-%d %H:%M:%S']:
            try: return datetime.strptime(value.strip(),fmt).date()
            except ValueError: pass
    return None

def status(value):
    n=norm(value)
    if not n: return 'Não informado'
    if n in ['aprovado','aprovada','regular']: return 'Aprovado'
    if n in ['ressalva','ressalvas','aprovado com ressalva','aprovado com ressalvas']: return 'Ressalva'
    if n in ['rejeitado','rejeitada','recusado','reprovado']: return 'Rejeitado'
    if n=='ausente': return 'Ausente'
    return 'Outro: '+str(value).strip()

def tipo(value):
    n=norm(value)
    return {'servico':'Serviço','servicos':'Serviço','consumo':'Consumo'}.get(n,'Outro / não informado')

def business_add(start,n,holidays):
    while n:
        start+=timedelta(days=1)
        if start.weekday()<5 and start not in holidays: n-=1
    return start

def business_roll(day,holidays):
    while day.weekday()>4 or day in holidays: day+=timedelta(days=1)
    return day

def brl(value):
    if value is None:return 'não disponível'
    return 'R$ '+f'{value:,.2f}'.replace(',','@').replace('.',',').replace('@','.')

def number(value):
    return 'não disponível' if value is None else f'{value:.1f}'.replace('.',',')

def stats(values):
    values=sorted(values)
    if not values:return {'n':0,'media':None,'mediana':None,'minimo':None,'maximo':None,'p90':None}
    return {'n':len(values),'media':statistics.mean(values),'mediana':statistics.median(values),'minimo':min(values),'maximo':max(values),'p90':values[math.ceil(.9*len(values))-1]}

def json_write(p,data):
    p.write_text(json.dumps(data,ensure_ascii=False,indent=2,default=lambda x:x.isoformat() if isinstance(x,(date,datetime)) else str(x)),encoding='utf-8')

def csv_write(p,rows):
    if not rows: rows=[{'resultado':'Sem registros elegíveis'}]
    with p.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=list(rows[0]),delimiter=';');w.writeheader()
        for row in rows:
            clean={}
            for k,v in row.items():
                if isinstance(v,(date,datetime)):v=v.isoformat()
                # Protege nomes e identificadores contra execução de fórmulas em CSV.
                if isinstance(v,str) and v.startswith(('=','+','-','@')):v="'"+v
                clean[k]=v
            w.writerow(clean)

def aggregates(rows,key):
    groups=defaultdict(list)
    for r in rows:groups[r[key]].append(r)
    result=[]
    for label,rs in groups.items():
        sums={m:sum(r[m] for r in rs if r[m] is not None) for m in MONEY}
        result.append({key:label,'registros':len(rs),'processos_unicos':len({r['processo'] for r in rs if r['processo']}),
          **{m:v/100 for m,v in sums.items()},
          **{m+'_preenchidos':sum(r[m] is not None for r in rs) for m in MONEY},
          'media_adiantado':statistics.mean(r['adiantado']/100 for r in rs if r['adiantado'] is not None) if any(r['adiantado'] is not None for r in rs) else None,
          'taxa_utilizacao':sums['utilizado']/sums['adiantado'] if sums['adiantado'] and all(r[m] is not None for r in rs for m in ['adiantado','utilizado']) else None,
          'cgm_aprovados':sum(r['cgm']=='Aprovado' for r in rs),'cgm_ressalvas':sum(r['cgm']=='Ressalva' for r in rs),'cgm_rejeitados':sum(r['cgm']=='Rejeitado' for r in rs)})
    return sorted(result,key=lambda r:(-r['adiantado'],str(r[key])))

def analyze(config,root):
    source=(root/config['arquivo']).resolve()
    wb=load_workbook(source,data_only=True,read_only=False)
    wf=load_workbook(source,data_only=False,read_only=False)
    ws=wb[config['aba']]; formulas=wf[config['aba']]
    begin=date.fromisoformat(config['inicio']); end=date.fromisoformat(config['fim']); observed=date.fromisoformat(config['data_observacao'])
    if end<begin or observed<end:raise ValueError('Período inválido ou observação anterior ao fim do período.')
    holidays={date.fromisoformat(x) for x in config.get('feriados',[])}
    issues=[]; rows=[]; excluded=[]; ignored=0
    def issue(row,code,message,field=''):
        issues.append({'linha_origem':row,'codigo':code,'campo':field,'descricao':message})
    # Fail-fast contra mudanças estruturais para evitar analisar a coluna errada.
    expected={'A':'processo','B':'servidor','C':'secretaria','D':'nota','E':'valor adt','F':'valor utl','G':'valor dev','I':'data pagto','W':'tipo','Z':'exame smf','AA':'exame cgm'}
    for col,text in expected.items():
        if text not in norm(ws[f'{col}2'].value):raise ValueError(f'Cabeçalho inesperado em {col}2: {ws[f"{col}2"].value!r}')
    for i in range(3,ws.max_row+1):
        raw={name:ws[f'{col}{i}'].value for col,name in COLS.items()}
        if not any(raw[k] is not None for k in ['processo','servidor','secretaria','nota','adiantado','utilizado','devolvido','pagamento']):
            ignored+=1;continue
        # Totais e textos isolados não são novos adiantamentos.
        if norm(raw['processo']).startswith('total'):
            excluded.append({'linha_origem':i,'motivo':'Linha de totalização'});continue
        r={'linha_origem':i,**raw}
        for k in ['processo','servidor','secretaria','nota']:r[k]=str(raw[k]).strip() if raw[k] is not None else ''
        r['servidor']=' '.join(r['servidor'].upper().split());r['secretaria']=' '.join(r['secretaria'].upper().split())
        for k in MONEY:
            r[k]=cents(raw[k])
            if r[k] is None:issue(i,'VALOR_AUSENTE_INVALIDO',f'Valor {k} ausente ou inválido.',k)
            elif r[k]<0:issue(i,'VALOR_NEGATIVO',f'Valor {k} negativo.',k)
        for k in DATES:r[k]=dt(raw[k])
        for col,name in COLS.items():
            fc=formulas[f'{col}{i}']
            if fc.data_type=='f' and raw[name] is None:issue(i,'FORMULA_SEM_CACHE','Fórmula sem resultado armazenado; recalcular a fonte no Excel.',name)
        if not r['processo']:issue(i,'PROCESSO_AUSENTE','Identificador de processo ausente.','processo')
        pay=r['pagamento']
        if not pay or not begin<=pay<=end:
            excluded.append({'linha_origem':i,'motivo':'Pagamento sem data válida' if not pay else 'Pagamento fora do período','processo':r['processo']})
            continue
        r['mes']=pay.strftime('%Y-%m')
        if norm(raw['mes_informado'])[:3]!=norm(MONTHS[pay.month-1]):issue(i,'MES_DIVERGENTE','Mês informado difere do mês do pagamento.','mes_informado')
        r['tipo_original']=str(raw['tipo'] or '').strip();r['tipo']=tipo(raw['tipo'])
        if r['tipo']=='Outro / não informado':issue(i,'TIPO_NAO_PADRONIZADO','Natureza diferente de serviço/consumo: '+r['tipo_original'],'tipo')
        r['smf']=status(raw['smf']);r['cgm']=status(raw['cgm'])
        for k in ['smf','cgm']:
            if r[k] in ['Não informado','Ausente'] or r[k].startswith('Outro:'):issue(i,'EXAME_AUSENTE_OUTRO',f'Exame {k.upper()}: {r[k]}.',k)
        if all(r[k] is not None for k in MONEY):
            r['diferenca_financeira']=r['adiantado']-r['utilizado']-r['devolvido']
            if r['diferenca_financeira']!=0:issue(i,'DIFERENCA_FINANCEIRA','Adiantado - utilizado - devolvido = '+brl(r['diferenca_financeira']/100)+'. Não corrigido automaticamente.')
        else:r['diferenca_financeira']=None
        if r['devolvido'] and norm(raw['devolucao']) in ['gasto total','sem uso']:issue(i,'DEVOLUCAO_CONTRADITORIA','Há valor devolvido, mas a data de devolução contém texto incompatível.','devolucao')
        if r['utilizado']==0 and r['ultima_aplicacao']:issue(i,'APLICACAO_SEM_VALOR','Última aplicação preenchida com valor utilizado igual a zero.','ultima_aplicacao')
        for k in DATES:
            if r[k] and r[k]>observed:issue(i,'DATA_FUTURA_OBSERVACAO',f'{k} posterior à data de observação.',k)
        for event,field in [('aplicacao','ultima_aplicacao'),('devolucao','devolucao'),('prestacao','prestacao')]:
            applicable=(r['utilizado'] is not None and r['utilizado']>0) if event=='aplicacao' else ((r['devolvido'] is not None and r['devolvido']>0) if event=='devolucao' else True)
            deadline=r['prazo_'+event]; actual=r[field]
            r['dias_pagamento_'+event]=None;r['atraso_calculado_'+event]=None
            if not applicable:
                state='Não aplicável' if (r['utilizado'] if event=='aplicacao' else r['devolvido'])==0 else 'Valor não informado'
            elif actual and pay<=actual<=observed:
                r['dias_pagamento_'+event]=(actual-pay).days
                if deadline and deadline<=observed:
                    r['atraso_calculado_'+event]=max(0,(actual-deadline).days)
                    state='Com atraso' if actual>deadline else 'No prazo'
                elif deadline:state='Vencimento após observação'
                else:state='Vencimento não informado'
            elif actual and actual<pay:
                state='Data anterior ao pagamento';issue(i,'EVENTO_ANTES_PAGAMENTO',f'{event} anterior ao pagamento.',field)
            elif actual and actual>observed:state='Evento após observação'
            else:
                state='Data não informada';issue(i,'DATA_EVENTO_AUSENTE',f'Data de {event} ausente ou inválida; não inferir inadimplência.',field)
            r['situacao_'+event]=state
            informed=norm(raw['atraso_informado_'+event]);calc=r['atraso_calculado_'+event]
            match=re.search(r'(\d+)\s*dia',informed)
            informed_days=int(match.group(1)) if match else (0 if informed=='nao' else None)
            if calc is not None and informed_days is not None and calc!=informed_days:issue(i,'ATRASO_DIVERGENTE',f'{event}: informado {informed_days} dias; calculado {calc} dias.','atraso_informado_'+event)
        # Comparação técnica: não substitui os prazos registrados e não julga exceções.
        expected_app=pay+timedelta(days=30)
        expected_ret=business_add(r['prazo_aplicacao'] or expected_app,3,holidays)
        expected_pres=business_roll((r['prazo_aplicacao'] or expected_app)+timedelta(days=10),holidays)
        for event,expected_day in [('aplicacao',expected_app),('devolucao',expected_ret),('prestacao',expected_pres)]:
            r['prazo_tecnico_'+event]=expected_day
            if r['prazo_'+event]!=expected_day:issue(i,'PRAZO_DIVERGENTE_'+event.upper(),f'Prazo registrado {r["prazo_"+event]}; referência técnica {expected_day}. Calendário validado: {config.get("calendario_validado",False)}.','prazo_'+event)
        rows.append(r)
    seen=defaultdict(list)
    for r in rows:seen[(r['processo'],r['nota'],r['tipo_original'])].append(r['linha_origem'])
    for key,lines in seen.items():
        if len(lines)>1:
            for i in lines:issue(i,'CHAVE_DUPLICADA','Mesma chave processo + nota + natureza nas linhas '+str(lines)+'. Registros mantidos para revisão.')
    monthly=[]
    first=begin.replace(day=1)
    while first<=end:
        label=first.strftime('%Y-%m');rs=[r for r in rows if r['mes']==label]
        a=aggregates(rs,'mes')
        monthly.append(a[0] if a else {'mes':label,'registros':0,'processos_unicos':0,**{m:0 for m in MONEY},**{m+'_preenchidos':0 for m in MONEY},'media_adiantado':None,'taxa_utilizacao':None,'cgm_aprovados':0,'cgm_ressalvas':0,'cgm_rejeitados':0})
        first=(first.replace(day=28)+timedelta(days=4)).replace(day=1)
    tables={'mensal':monthly,'secretarias':aggregates(rows,'secretaria'),'servidores':aggregates(rows,'servidor'),'naturezas':aggregates(rows,'tipo')}
    # Consolidação por mês E secretaria para conciliação e uso no RCI.
    sec_month=[]
    for sec in sorted({r['secretaria'] for r in rows}):
        for m in monthly:
            rs=[r for r in rows if r['secretaria']==sec and r['mes']==m['mes']]
            sec_month.append({'secretaria':sec,'mes':m['mes'],'registros':len(rs),**{k:sum(r[k] for r in rs if r[k] is not None)/100 for k in MONEY}})
    tables['secretaria_mes']=sec_month
    exams=[]
    for organ in ['smf','cgm']:
        for label in sorted({r[organ] for r in rows}):
            rs=[r for r in rows if r[organ]==label]
            exams.append({'orgao':organ.upper(),'resultado':label,'registros':len(rs),'percentual':len(rs)/len(rows) if rows else None,**{m:sum(r[m] for r in rs if r[m] is not None)/100 for m in MONEY}})
    tables['exames']=exams
    cross=[]
    for (smf,cgm),n in sorted(Counter((r['smf'],r['cgm']) for r in rows).items()):cross.append({'smf':smf,'cgm':cgm,'registros':n,'percentual':n/len(rows) if rows else None})
    tables['cruzamento_exames']=cross
    temporal=[];timing_cases=[]
    for event in ['aplicacao','devolucao','prestacao']:
        metric='dias_pagamento_'+event;lag='atraso_calculado_'+event
        durations=stats([r[metric] for r in rows if r[metric] is not None])
        delays=stats([r[lag] for r in rows if r[lag] is not None and r[lag]>0])
        valid=sum(r[lag] is not None for r in rows);late=sum(r[lag] is not None and r[lag]>0 for r in rows)
        temporal.append({'evento':event,**durations,'avaliaveis_prazo':valid,'com_atraso':late,'taxa_atraso':late/valid if valid else None,'media_atraso_apenas_atrasados':delays['media'],'maior_atraso':delays['maximo']})
        for r in rows:
            timing_cases.append({'evento':event,**{k:r[k] for k in ['linha_origem','processo','nota','servidor','secretaria']},'dias_desde_pagamento':r[metric],'atraso_dias':r[lag],'situacao':r['situacao_'+event]})
    tables['tempos']=temporal;tables['casos_prazos']=timing_cases
    extremes=[]
    for event in ['aplicacao','devolucao','prestacao']:
        eligible=[r for r in timing_cases if r['evento']==event and r['dias_desde_pagamento'] is not None]
        if not eligible:continue
        for criterion,field,fn in [('Menor duração','dias_desde_pagamento',min),('Maior duração','dias_desde_pagamento',max),('Maior atraso registrado','atraso_dias',max)]:
            valid=[r[field] for r in eligible if r[field] is not None]
            if not valid:continue
            v=fn(valid)
            if criterion=='Maior atraso registrado' and not v:continue
            for r in eligible:
                if r[field]==v:extremes.append({'evento':event,'criterio':criterion,'dias':v,**{k:r[k] for k in ['linha_origem','processo','nota','servidor','secretaria']}})
    tables['extremos_prazos']=extremes
    grouped_times=[]
    for key in ['secretaria','servidor']:
        for label in sorted({r[key] for r in rows}):
            for event in ['aplicacao','devolucao','prestacao']:
                rs=[r for r in rows if r[key]==label]
                durations=stats([r['dias_pagamento_'+event] for r in rs if r['dias_pagamento_'+event] is not None])
                evaluable=[r['atraso_calculado_'+event] for r in rs if r['atraso_calculado_'+event] is not None]
                grouped_times.append({'agrupamento':key,'grupo':label,'evento':event,**durations,'avaliaveis_prazo':len(evaluable),'com_atraso':sum(x>0 for x in evaluable),'maior_atraso':max(evaluable) if evaluable else None})
    tables['tempos_por_grupo']=grouped_times
    financial=stats([r['adiantado']/100 for r in rows if r['adiantado'] is not None])
    total={m:sum(r[m] for r in rows if r[m] is not None)/100 for m in MONEY}
    summary={'registros':len(rows),'processos_unicos':len({r['processo'] for r in rows if r['processo']}),'servidores':len({r['servidor'] for r in rows if r['servidor']}),'secretarias':len({r['secretaria'] for r in rows if r['secretaria']}),**total,'diferenca_financeira':sum(r['diferenca_financeira'] for r in rows if r['diferenca_financeira'] is not None)/100,'adiantado_estatisticas':financial,'aprovados_ambos':sum(r['smf']==r['cgm']=='Aprovado' for r in rows),'exames_divergentes':sum(r['smf']!=r['cgm'] for r in rows),'sem_uso':sum(r['utilizado']==0 for r in rows),'devolucao_integral':sum(r['adiantado'] is not None and r['adiantado']>0 and r['utilizado']==0 and r['devolvido']==r['adiantado'] for r in rows),'gasto_integral':sum(r['adiantado'] is not None and r['utilizado']==r['adiantado'] and r['devolvido']==0 for r in rows),'eventos_pos_semestre':{k:sum(r[k] is not None and r[k]>end for r in rows) for k in ['ultima_aplicacao','devolucao','prestacao']},'linhas_ignoradas':ignored,'linhas_excluidas':len(excluded),'alertas':len(issues),'linhas_com_alerta':len({x['linha_origem'] for x in issues if isinstance(x['linha_origem'],int)}),'valores_ausentes':{m:sum(r[m] is None for r in rows) for m in MONEY}}
    # Concilia caches da aba consolidada com os registros, sem recalcular a fonte.
    reconciliation=[]
    if config.get('aba_consolidado') in wb.sheetnames:
        cs=wb[config['aba_consolidado']]
        for i in range(3,cs.max_row+1):
            key=str(cs.cell(i,1).value or '').strip().upper()
            is_total=bool(re.search(r'total',norm(key))) or (i==24 and not key)
            if not key and not is_total:continue
            for month in monthly:
                mi=int(month['mes'][-2:]); base=3+(mi-1)*3
                if not 1<=mi<=6:continue
                for off,metric in enumerate(MONEY):
                    cached=cents(cs.cell(i,base+off).value)
                    relevant=[r for r in rows if r['mes']==month['mes'] and (is_total or r['secretaria']==key)]
                    actual=sum(r[metric] for r in relevant if r[metric] is not None)
                    reconciliation.append({'secretaria':'TOTAL' if is_total else key,'mes':month['mes'],'metrica':metric,'celula_origem':cs.cell(i,base+off).coordinate,'consolidado':cached/100 if cached is not None else None,'recalculado':actual/100,'diferenca':(actual-cached)/100 if cached is not None else None})
            for col,metric in zip([24,25,26],MONEY):
                cached=cents(cs.cell(i,col).value)
                relevant=[r for r in rows if is_total or r['secretaria']==key]
                actual=sum(r[metric] for r in relevant if r[metric] is not None)
                reconciliation.append({'secretaria':'TOTAL' if is_total else key,'mes':'PERIODO','metrica':metric,'celula_origem':cs.cell(i,col).coordinate,'consolidado':cached/100 if cached is not None else None,'recalculado':actual/100,'diferenca':(actual-cached)/100 if cached is not None else None})
        known={str(cs.cell(i,1).value or '').strip().upper() for i in range(3,cs.max_row+1)}
        for sec in {r['secretaria'] for r in rows}-known:issue('consolidado','SECRETARIA_SEM_CONSOLIDADO','Secretaria não localizada no consolidado: '+sec)
    tables['conciliacao']=reconciliation
    summary['conciliacao_divergencias']=sum(r['diferenca'] not in (0,None) for r in reconciliation)
    summary['conciliacao_sem_cache']=sum(r['consolidado'] is None for r in reconciliation)
    summary['alertas']=len(issues)
    tables['alertas']=issues;tables['excluidos']=excluded
    return rows,tables,summary,{'arquivo':source.name,'sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'aba':ws.title,'versao':VERSION,'config':config,'cabecalho_linha':2,'unidade':'registro por nota/natureza; processo SEI pode repetir','valores':'centavos inteiros nos cálculos; valores monetários em reais nas saídas','consolidado':'Resultados armazenados no XLSX, confrontados sem alterar a fonte.'}

def narratives(rows,tables,s,cfg):
    sections=[];n=s['registros'];total=s['adiantado']
    pct=lambda a,b: f'{100*a/b:.2f}%'.replace('.',',') if b else 'não disponível'
    text=(f"Foram identificados {n} registros por nota/natureza, vinculados a {s['processos_unicos']} processos SEI únicos, {s['servidores']} servidores e {s['secretarias']} secretarias. "
      f"O recorte abrange pagamentos de {date.fromisoformat(cfg['inicio']).strftime('%d/%m/%Y')} a {date.fromisoformat(cfg['fim']).strftime('%d/%m/%Y')}, com informações disponíveis na planilha observadas em {date.fromisoformat(cfg['data_observacao']).strftime('%d/%m/%Y')}. "
      f"Eventos posteriores ao semestre foram mantidos: {s['eventos_pos_semestre']['prestacao']} datas de prestação e {s['eventos_pos_semestre']['devolucao']} datas de devolução. Não se trata de uma fotografia dos processos encerrados em 30 de junho. "
      f"A base registra {brl(total)} em adiantamentos, {brl(s['utilizado'])} utilizados ({pct(s['utilizado'],total)}) e {brl(s['devolvido'])} devolvidos ({pct(s['devolvido'],total)}). "
      f"A diferença agregada adiantado menos utilizado e devolvido é de {brl(s['diferenca_financeira'])}. Há {sum(s['valores_ausentes'].values())} campos monetários ausentes. Os totais reproduzem a fonte e exigem a conferência dos alertas antes da versão definitiva.")
    if tables['mensal']:
        peak=max(tables['mensal'],key=lambda r:r['adiantado']);low=min(tables['mensal'],key=lambda r:r['adiantado'])
        text+=f" O maior total mensal adiantado foi de {brl(peak['adiantado'])} em {peak['mes']}, e o menor, de {brl(low['adiantado'])} em {low['mes']}."
    sections.append(('01_panorama','Panorama e conciliação',text,'mensal','mensal'))
    sec=tables['secretarias'];ser=tables['servidores'];financial=s['adiantado_estatisticas']
    def tied(table,key,metric,mode='max'):
        vals=[r[metric] for r in table if r[metric] is not None]
        if not vals:return 'não disponível'
        v=(max(vals) if mode=='max' else min(vals));names=[str(r[key]) for r in table if r[metric]==v]
        return ', '.join(names)+f' ({brl(v) if metric in MONEY else v})'
    text=(f"Entre as secretarias presentes na base, o maior valor adiantado pertence a {tied(sec,'secretaria','adiantado')}, e o maior valor utilizado a {tied(sec,'secretaria','utilizado')}. "
      f"A maior quantidade de registros foi de {tied(sec,'secretaria','registros')}. Entre servidores, o maior total adiantado foi de {tied(ser,'servidor','adiantado')}, o maior utilizado de {tied(ser,'servidor','utilizado')} e a maior quantidade de registros de {tied(ser,'servidor','registros')}. "
      f"A média por registro de adiantamento foi de {brl(financial['media'])}, mediana de {brl(financial['mediana'])}, mínimo de {brl(financial['minimo'])} e máximo de {brl(financial['maximo'])}, para {financial['n']} valores válidos. "
      f"A menor soma utilizada entre as secretarias com registros foi de {tied(sec,'secretaria','utilizado','min')}. Os rankings completos constam das tabelas. Quantidade de registros não equivale a quantidade de solicitações: um mesmo pedido/processo pode conter notas de consumo e serviço. Concentração financeira, isoladamente, não comprova irregularidade.")
    sections.append(('02_distribuicao','Distribuição por secretaria e servidor',text,'secretarias','secretarias'))
    text=(f"A classificação original foi mantida e normalizada apenas em grafia e espaços. A distribuição por natureza consta da tabela. "
      f"Há {s['sem_uso']} registros com valor utilizado igual a zero, {s['devolucao_integral']} com devolução integral aritmeticamente compatível e {s['gasto_integral']} com utilização integral registrada. "
      "As taxas de utilização são ponderadas pelos valores adiantados, e não médias simples dos percentuais individuais. Uma devolução elevada pode orientar a revisão das estimativas, mas não demonstra, por si só, falta de planejamento. Classificações diferentes de serviço/consumo ficam em grupo próprio para conferência.")
    sections.append(('03_naturezas','Natureza e utilização dos recursos',text,'naturezas','naturezas'))
    paragraphs=[]
    for t in tables['tempos']:
        label={'aplicacao':'última aplicação','devolucao':'devolução','prestacao':'prestação de contas'}[t['evento']]
        paragraphs.append(f"Do pagamento até a {label}, a duração média observada foi de {number(t['media'])} dias, mediana de {number(t['mediana'])}, mínimo de {number(t['minimo'])} e máximo de {number(t['maximo'])}, com {t['n']} datas elegíveis. Em relação ao vencimento registrado na planilha, {t['com_atraso']} de {t['avaliaveis_prazo']} registros avaliáveis tiveram atraso ({pct(t['com_atraso'],t['avaliaveis_prazo'])}); o maior atraso foi de {number(t['maior_atraso'])} dias.")
    paragraphs.append('As diferenças de datas usam dias corridos. Não há data da primeira despesa na base: a duração de aplicação mede a última utilização, sem permitir afirmar quando o servidor começou a usar o recurso. A data de pagamento foi usada como aproximação da disponibilização do numerário. A data de prestação é a informada na planilha, sem comprovação do protocolo de recebimento na Fazenda. Datas ausentes ou contraditórias não entram nas médias e não foram tratadas como zero. Os resultados de atraso são descritivos dos vencimentos informados, sujeitos à revisão normativa e do calendário.')
    extreme_text=[]
    for event in ['devolucao','prestacao']:
        fast=[r for r in tables['extremos_prazos'] if r['evento']==event and r['criterio']=='Menor duração']
        if fast:extreme_text.append(f"Menor duração até {event}: "+', '.join(sorted({r['servidor'] for r in fast}))+f" ({fast[0]['dias']} dias).")
    for event in ['aplicacao','devolucao','prestacao']:
        late=[r for r in tables['extremos_prazos'] if r['evento']==event and r['criterio']=='Maior atraso registrado']
        if late:extreme_text.append(f"Maior atraso de {event}: "+', '.join(sorted({r['servidor'] for r in late}))+f" ({late[0]['dias']} dias).")
    paragraphs.append(' '.join(extreme_text)+' A relação completa de extremos e empates, inclusive última aplicação no mesmo dia do pagamento, está em extremos_prazos.csv. Tempos por secretaria e servidor constam de tempos_por_grupo.csv, com o tamanho de cada amostra.')
    sections.append(('04_prazos','Durações e atrasos registrados','\n\n'.join(paragraphs),'tempos','prazos'))
    organtexts=[]
    for organ in ['SMF','CGM']:
        organtexts.append(organ+': '+', '.join(f"{r['resultado']} {r['registros']} ({pct(r['registros'],n)})" for r in tables['exames'] if r['orgao']==organ)+'.')
    text=' '.join(organtexts)+f" Houve {s['aprovados_ambos']} registros aprovados simultaneamente por SMF e CGM e {s['exames_divergentes']} registros com categorias diferentes entre os órgãos, incluindo ausências. Os exames são os resultados registrados para a prestação de contas e não representam a aprovação inicial do pedido de adiantamento. Ressalva não foi contada como aprovação sem ressalva. A tabela cruzada preserva todas as combinações e a tabela por órgão informa os montantes envolvidos; rejeição registrada não equivale automaticamente a dano ou valor a ressarcir."
    sections.append(('05_exames','Resultados dos exames SMF e CGM',text,'exames','exames'))
    codes=Counter(x['codigo'] for x in tables['alertas'])
    text=(f"Foram produzidos {s['alertas']} alertas, com {s['linhas_com_alerta']} linhas de dados sinalizadas. A conciliação com os valores armazenados na aba por secretaria identificou {s['conciliacao_divergencias']} diferenças de células e {s['conciliacao_sem_cache']} valores sem cache, sem sobrescrever a fonte. "
      "Esses alertas identificam campos para revisão, sem concluir automaticamente pela existência de infração. Os códigos e as linhas estão em alertas.csv. "
      "A IN-001/2018-CGM retificada em 2025, no art. 13 (página 5 do PDF fornecido), prevê três dias úteis para devolução. Os arts. 6 e 17 e o §3 do art. 17 (páginas 3, 6 e 7) fundamentam a comparação técnica de aplicação e prestação. "
      f"O calendário municipal está marcado como {'validado' if cfg.get('calendario_validado') else 'NÃO validado'}; sem validação, a comparação técnica considera fins de semana e apenas os feriados explicitamente configurados. Não substitui os vencimentos da planilha nem contempla autorizações excepcionais, cancelamentos e fechamento do exercício. "
      "O modelo de RCI foi utilizado como referência de organização (seção F.4, páginas 43 a 57), sem reaproveitar números ou conclusões de 2025. OP, empenho, observações e colunas auxiliares não foram usados para gerar inferências. Antes de anexar a versão definitiva, conferir a fonte, os valores, o calendário e os pareceres com Katryanne.")
    sections.append(('06_revisao','Qualidade dos dados e critérios de revisão',text,'alertas',None))
    return sections

def charts(tables,path):
    os.environ.setdefault('MPLCONFIGDIR',str(Path(__file__).resolve().parents[1]/'work/matplotlib'))
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
    plt.rcParams.update({'font.family':'DejaVu Sans','font.size':10,'axes.spines.top':False,'axes.spines.right':False})
    palette=['#24476B','#3B8686','#D7983B']
    def save(fig,name):
        fig.tight_layout();fig.savefig(path/(name+'.png'),dpi=190,bbox_inches='tight');fig.savefig(path/(name+'.svg'),bbox_inches='tight');plt.close(fig)
    formatter=FuncFormatter(lambda x,pos:f'{x:,.0f}'.replace(',','.'))
    rs=tables['mensal'];fig,ax=plt.subplots(figsize=(10,4.4));x=list(range(len(rs)))
    for j,k in enumerate(MONEY):ax.bar([v+(j-1)*.25 for v in x],[r[k] for r in rs],width=.25,label={'adiantado':'Adiantado','utilizado':'Utilizado','devolvido':'Devolvido'}[k],color=palette[j])
    ax.set_xticks(x,[r['mes'] for r in rs]);ax.yaxis.set_major_formatter(formatter);ax.set_ylabel('Valores registrados (R$)');ax.legend(ncol=3,frameon=False);ax.set_title('Adiantamentos por mês do pagamento');save(fig,'mensal')
    rs=sorted(tables['secretarias'],key=lambda r:r['adiantado'])
    fig,ax=plt.subplots(figsize=(10,max(4,len(rs)*.3)))
    ax.barh([r['secretaria'] for r in rs],[r['adiantado'] for r in rs],color=palette[0]);ax.set_xlabel('Valor adiantado (R$)');ax.xaxis.set_major_formatter(formatter);ax.set_title('Valores adiantados por secretaria');save(fig,'secretarias')
    rs=tables['naturezas'];fig,ax=plt.subplots(figsize=(10,4.3));x=list(range(len(rs)))
    labels=['Outro' if r['tipo'].startswith('Outro') else r['tipo'] for r in rs]
    for j,k in enumerate(MONEY):ax.bar([v+(j-1)*.25 for v in x],[r[k] for r in rs],width=.25,label=k.capitalize(),color=palette[j])
    ax.set_xticks(x,labels);ax.yaxis.set_major_formatter(formatter);ax.set_ylabel('Valores registrados (R$)');ax.legend(frameon=False,ncol=3);ax.set_title('Distribuição por natureza');save(fig,'naturezas')
    rs=tables['tempos'];fig,ax=plt.subplots(figsize=(10,4.3));ax.bar(['Última aplicação','Devolução','Prestação'],[r['media'] or 0 for r in rs],color=palette)
    for i,r in enumerate(rs):ax.text(i,(r['media'] or 0)+.5,f"{number(r['media'])} dias | n={r['n']}",ha='center')
    ax.set_ylim(0,max([r['media'] or 0 for r in rs]+[1])*1.22);ax.set_ylabel('Dias corridos após pagamento');ax.set_title('Duração média para eventos com datas elegíveis');save(fig,'prazos')
    cats=['Aprovado','Ressalva','Rejeitado','Ausente','Não informado'];fig,ax=plt.subplots(figsize=(10,4.3));x=list(range(len(cats)))
    for j,o in enumerate(['SMF','CGM']):
        vals=[sum(r['registros'] for r in tables['exames'] if r['orgao']==o and r['resultado']==c) for c in cats]
        ax.bar([v+(j-.5)*.35 for v in x],vals,.35,label=o,color=palette[j])
    ax.set_xticks(x,cats);ax.set_ylabel('Registros');ax.legend(frameon=False);ax.set_title('Exames das prestações de contas');save(fig,'exames')

def create_documents(run,sections,summary,tables,meta):
    from docx import Document
    from docx.shared import Inches,Pt,RGBColor
    from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Image,PageBreak,Table,TableStyle
    from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.utils import ImageReader
    from xml.sax.saxutils import escape
    cfg=meta['config'];label='PRELIMINAR - TESTE' if cfg['preliminar'] else 'VERSÃO PARA REVISÃO'
    styles=getSampleStyleSheet();styles.add(ParagraphStyle(name='BodyPT',fontName='Helvetica',fontSize=10.2,leading=14,spaceAfter=9));styles['Title'].fontSize=19;styles['Title'].leading=23
    def doc_new():
        d=Document();sec=d.sections[0];sec.page_width=Inches(8.5);sec.page_height=Inches(11);sec.top_margin=sec.bottom_margin=Inches(.7);sec.left_margin=sec.right_margin=Inches(.8)
        for name in ['Normal','Title','Heading 1','Heading 2']:
            d.styles[name].font.name='Arial';d.styles[name].font.color.rgb=RGBColor(0,0,0)
        d.styles['Normal'].font.size=Pt(10.5);d.styles['Normal'].paragraph_format.space_after=Pt(8)
        d.core_properties.author='';d.core_properties.title='Adiantamentos - '+label
        return d
    def p(text,style='BodyPT'):return Paragraph(escape(text),styles[style])
    def footer(canvas,doc):
        canvas.setFont('Helvetica',8);canvas.setFillColor(colors.HexColor('#555555'));canvas.drawString(50,25,label+' | Adiantamentos | '+cfg['inicio']+' a '+cfg['fim']);canvas.drawRightString(560,25,str(doc.page))
    def render_pdf(path,story):
        SimpleDocTemplate(str(path),pagesize=(612,792),leftMargin=56,rightMargin=56,topMargin=42,bottomMargin=42,title='Adiantamentos - '+label,author='').build(story,onFirstPage=footer,onLaterPages=footer)
    doc=doc_new();story=[]
    for idx,(folder,title,text,table_name,chart) in enumerate(sections):
        part=run/folder;part.mkdir(exist_ok=True)
        (part/'texto_para_RCI.txt').write_text(label+'\n'+title+'\n\n'+text+'\n',encoding='utf-8')
        csv_write(part/(table_name+'.csv'),tables[table_name])
        if idx:doc.add_page_break();story.append(PageBreak())
        doc.add_heading(title,0);doc.add_paragraph(label)
        story.extend([p(title,'Title'),p(label),Spacer(1,8)])
        for para in text.split('\n\n'):
            doc.add_paragraph(para);story.append(p(para))
        if chart:
            pic=run/'graficos'/(chart+'.png');doc.add_picture(str(pic),width=Inches(6.5))
            iw,ih=ImageReader(str(pic)).getSize();h=min(240,ih*500/iw);w=h*iw/ih
            story.extend([Spacer(1,7),Image(str(pic),width=w,height=h)])
        caption='Fonte: '+meta['arquivo']+', aba '+cfg['aba']+'. Pagamentos no período; situação observada em '+cfg['data_observacao']+'.'
        doc.add_paragraph(caption);story.append(p(caption))
    doc.save(run/'Relatorio_Adiantamentos_PRELIMINAR.docx');render_pdf(run/'Relatorio_Adiantamentos_PRELIMINAR.pdf',story)
    guide=doc_new();guide.add_heading('Guia dos arquivos gerados',0);guide.add_paragraph(label)
    story=[p('Guia dos arquivos gerados','Title'),p(label)]
    paragraphs=[
      'O pacote reúne resultados, gráficos e textos editáveis para a seção de adiantamentos do RCI. Comece pelo relatório geral, confira os alertas e selecione os trechos que deseja inserir manualmente. Esta execução não modifica o Excel original.',
      f"Período dos pagamentos: {cfg['inicio']} a {cfg['fim']}. Observação: {cfg['data_observacao']}. Base: {summary['registros']} registros e {summary['processos_unicos']} processos SEI. Cada execução cria pasta própria para preservar correções manuais das versões anteriores.",
      'Relatorio_Adiantamentos_PRELIMINAR.docx e .pdf: apresentam a análise em seis partes, com textos e gráficos. O Word é editável; o PDF oferece uma versão de leitura.',
      'Planilhas/Analise_editavel.xlsx: reúne as tabelas de apoio, indicadores e base normalizada com linha de origem. Os resultados são um retrato da execução: para atualizar cálculos e textos, corrija a fonte ou configuração e rode novamente. Alterações manuais nesta planilha não recalculam automaticamente o relatório.',
      'graficos/: cinco figuras em PNG para inserção no RCI e em SVG vetorial para edição. Cada figura tem seu texto explicativo na pasta temática correspondente.',
      'dados/: base normalizada em CSV/JSON, indicadores, tabelas completas, casos de prazo com processos e notas, exclusões, conciliação do consolidado, alertas e resumo agregado para IA. Os CSV usam ponto e vírgula e UTF-8 com BOM; importe números usando ponto decimal.',
      'manifesto.json: configuração, versão, nomes e hashes SHA-256 das fontes e dos arquivos gerados. Permite saber exatamente qual planilha originou o resultado. Alterações manuais posteriores deixam de corresponder aos hashes registrados.',
      'textos/: textos reunidos e, quando solicitado, uma abertura e encerramento do Ollama combinados com o corpo determinístico. O modelo não recebe dados reais e não redige os números. O texto continua sujeito à revisão.',
      'As pastas 01 a 06 contêm o texto específico para copiar e uma tabela de apoio. As tabelas completas adicionais ficam em dados/ e na planilha editável.',
      'Revisão: conferir diferenças financeiras, natureza fora do padrão, datas contraditórias, vencimentos, resultados dos pareceres e conciliação. Validar o calendário de dias úteis antes de qualquer conclusão normativa. Não classificar automaticamente servidores em alcance nem inferir dano.',
      'Nova execução: atualizar o arquivo da fonte, copiar config.exemplo.json para config.local.json, ajustar período/aba/data de observação/saída e executar analisar.ps1. Nunca é necessário apagar relatórios anteriores. ULTIMA_EXECUCAO.txt informa a pasta mais recente.'
    ]
    for idx,txt in enumerate(paragraphs):
        if idx==6:guide.add_page_break();story.append(PageBreak())
        guide.add_paragraph(txt);story.append(p(txt))
    for folder,title,*_ in sections:
        txt=folder+'/: '+title+'. Contém a tabela e o texto correspondente.';guide.add_paragraph(txt);story.append(p(txt))
    guide.add_page_break();guide.add_heading('Inventário dos arquivos',1)
    story.extend([PageBreak(),p('Inventário dos arquivos','Heading1')])
    purposes={'mensal':'Evolução dos valores e registros por mês de pagamento.','secretarias':'Totais e frequência por secretaria.','servidores':'Ranking completo de valores e registros por servidor.','naturezas':'Distribuição entre serviço, consumo e outras naturezas.','secretaria_mes':'Valores por secretaria e mês.','exames':'Quantidade, percentual e valores por resultado de cada órgão.','cruzamento_exames':'Combinações entre resultados SMF e CGM.','tempos':'Média, mediana, extremos e cobertura das durações e atrasos.','casos_prazos':'Datas avaliáveis e atrasos com identificação da linha, nota e processo.','extremos_prazos':'Menores e maiores durações e maiores atrasos, incluindo empates.','tempos_por_grupo':'Estatísticas de duração e atraso por secretaria e servidor.','conciliacao':'Confronto célula a célula com o consolidado da fonte.','alertas':'Campos e critérios que precisam de conferência, com linha original.','excluidos':'Linhas excluídas do recorte e respectivos motivos.','base_normalizada':'Registros normalizados para rastreabilidade e reutilização.','indicadores':'Resumo numérico legível por programas.','resumo_ia':'Resumo agregado local, sem identificadores pessoais. Não enviado ao modelo.','execucao_ia':'Modelo, duração e escopo da geração opcional.','texto_para_RCI':'Texto específico da seção, editável e pronto para seleção manual.','textos_para_RCI':'Todos os textos determinísticos reunidos.','rascunho_ollama_REVISAR':'Abertura e encerramento opcionais combinados com o corpo verificado.','manifesto':'Fontes, configuração, versão e hashes dos arquivos.','Analise_editavel':'Planilha com tabelas, indicadores e base para consulta e edição.','Relatorio_Adiantamentos_PRELIMINAR':'Relatório geral com seis partes.','LEIA_PRIMEIRO_Guia_dos_arquivos':'Guia de uso e inventário dos arquivos.'}
    names={str(p.relative_to(run)).replace('\\','/') for p in run.rglob('*') if p.is_file()}
    names.update(['LEIA_PRIMEIRO_Guia_dos_arquivos.docx','LEIA_PRIMEIRO_Guia_dos_arquivos.pdf','manifesto.json'])
    catalog=[{'arquivo':'INVENTARIO_ARQUIVOS.csv','finalidade':'Lista editável dos arquivos e suas finalidades.'}]
    for name in sorted(names):
        stem=Path(name).stem;ext=Path(name).suffix
        why=purposes.get(stem,'Material de apoio à análise.')
        if name.startswith('graficos/'):
            why='Figura para inserção no RCI. '+('PNG para colagem.' if ext=='.png' else 'SVG vetorial para edição.')
        catalog.append({'arquivo':name,'finalidade':why})
    catalog.append({'arquivo':'dados/execucao_ia.json e textos/rascunho_ollama_REVISAR.txt (opcionais)','finalidade':'Criados apenas se a geração local com Ollama for executada com sucesso.'})
    csv_write(run/'INVENTARIO_ARQUIVOS.csv',catalog)
    guide.add_paragraph('INVENTARIO_ARQUIVOS.csv reproduz a lista abaixo em formato editável. Arquivos opcionais de IA são acrescentados após a execução principal. Os PDFs foram verificados visualmente; a paginação dos DOCX deve ser conferida no Word, pois a renderização local de Word não ficou disponível nesta sessão.')
    story.append(p('INVENTARIO_ARQUIVOS.csv reproduz a lista abaixo em formato editável. Arquivos opcionais de IA são acrescentados após a execução principal. Os PDFs foram verificados visualmente; a paginação dos DOCX deve ser conferida no Word, pois a renderização local de Word não ficou disponível nesta sessão.'))
    table=guide.add_table(rows=1,cols=2);table.style='Light Shading Accent 1';table.rows[0].cells[0].text='Arquivo';table.rows[0].cells[1].text='Finalidade'
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    for item in catalog:
        cells=table.add_row().cells;cells[0].text=item['arquivo'];cells[1].text=item['finalidade']
    for row in table.rows:
        for j,cell in enumerate(row.cells):
            cell.width=Inches(3.25)
            props=cell._tc.get_or_add_tcPr();borders=OxmlElement('w:tcBorders')
            for edge in ['top','left','bottom','right']:
                el=OxmlElement('w:'+edge);el.set(qn('w:val'),'single');el.set(qn('w:sz'),'4');el.set(qn('w:color'),'D9D9D9');borders.append(el)
            props.append(borders)
            for para in cell.paragraphs:
                for r in para.runs:r.font.size=Pt(9)
    grid=[[p('Arquivo'),p('Finalidade')]]+[[p(x['arquivo']),p(x['finalidade'])] for x in catalog]
    pt=Table(grid,colWidths=[250,250],repeatRows=1,hAlign='LEFT');pt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#DDE6EF')),('GRID',(0,0),(-1,-1),.4,colors.HexColor('#D9D9D9')),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),6),('RIGHTPADDING',(0,0),(-1,-1),6),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]));story.append(pt)
    guide.save(run/'LEIA_PRIMEIRO_Guia_dos_arquivos.docx');render_pdf(run/'LEIA_PRIMEIRO_Guia_dos_arquivos.pdf',story)

def export_workbook(root,run,tables,rows,summary):
    payload={'tables':tables,'rows':[{k:(v/100 if k in MONEY or k=='diferenca_financeira' else v) for k,v in r.items()} for r in rows], 'summary':summary}
    json_write(run/'dados/workbook_payload.json',payload)
    node=os.environ.get('ADIANTAMENTO_NODE')
    if node:
        subprocess.run([node,str(root/'src/exportar_planilha.mjs'),str(run)],check=True)
        support=root/'work'/run.name;support.mkdir(parents=True,exist_ok=True)
        (run/'dados/workbook_payload.json').rename(support/'workbook_payload.json')
        return
    # Ambiente portátil sem artifact-tool: fallback explícito. Nunca edita a fonte.
    from openpyxl import Workbook
    from openpyxl.styles import Font,PatternFill,Alignment
    from openpyxl.utils import get_column_letter
    wb=Workbook();wb.remove(wb.active)
    collections={'Resumo':[{'indicador':k,'valor':v} for k,v in summary.items() if not isinstance(v,dict)],**tables,'Base normalizada':payload['rows']}
    for name,rs in collections.items():
        ws=wb.create_sheet(name[:31]);ws.sheet_view.showGridLines=False
        if not rs:rs=[{'resultado':'Sem registros'}]
        headers=list(rs[0]);ws.append(headers)
        for row in rs:
            ws.append([("'"+v if isinstance(v,str) and v.startswith(('=','+','-','@')) else v) for v in row.values()])
        for c in ws[1]:c.fill=PatternFill('solid',fgColor='24476B');c.font=Font(color='FFFFFF',bold=True);c.alignment=Alignment(wrap_text=True)
        for i,h in enumerate(headers,1):
            ws.column_dimensions[get_column_letter(i)].width=min(55,max(18,len(h)+2))
            if h in MONEY or h.startswith('media_adiantado') or h=='diferenca_financeira':
                for cells in ws.iter_cols(min_col=i,max_col=i,min_row=2):
                    for cell in cells:cell.number_format='"R$" #,##0.00'
        ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    wb.save(run/'Planilhas/Analise_editavel.xlsx')
    support=root/'work'/run.name;support.mkdir(parents=True,exist_ok=True)
    (run/'dados/workbook_payload.json').rename(support/'workbook_payload.json')

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--config',default='config.exemplo.json');args=ap.parse_args()
    root=Path(__file__).resolve().parents[1];config=json.loads((root/args.config).read_text(encoding='utf-8-sig'))
    rows,tables,summary,meta=analyze(config,root)
    if not rows:raise SystemExit('Nenhum registro elegível: confira período, aba e datas. Nenhum relatório foi gerado.')
    out=(root/config['saida']).resolve();out.mkdir(parents=True,exist_ok=True)
    stamp=datetime.now().strftime('%Y%m%d-%H%M%S-%f');run=out/('TESTE_'+stamp if config['preliminar'] else 'REVISAO_'+stamp);run.mkdir()
    for part in ['dados','graficos','textos','Planilhas']:(run/part).mkdir()
    for name,data in tables.items():csv_write(run/'dados'/(name+'.csv'),data)
    normalized=[{k:(v/100 if k in MONEY or k=='diferenca_financeira' else v) for k,v in r.items()} for r in rows]
    csv_write(run/'dados/base_normalizada.csv',normalized);json_write(run/'dados/base_normalizada.json',normalized);json_write(run/'dados/indicadores.json',summary)
    sections=narratives(rows,tables,summary,config)
    (run/'textos/textos_para_RCI.txt').write_text('\n\n'.join(title+'\n'+text for _,title,text,*_ in sections),encoding='utf-8')
    json_write(run/'dados/resumo_ia.json',{'periodo':{'inicio':config['inicio'],'fim':config['fim'],'observacao':config['data_observacao'],'preliminar':config['preliminar']},'indicadores':summary,'mensal':tables['mensal'],'naturezas':tables['naturezas'],'exames':tables['exames'],'prazos':tables['tempos'],'limitacoes':['Atraso segundo vencimento informado; calendário não validado' if not config.get('calendario_validado') else 'Calendário informado pelo operador','Apenas última aplicação disponível','Sem comprovação de data de protocolo SMF','Base não comprovadamente completa']})
    charts(tables,run/'graficos');export_workbook(root,run,tables,rows,summary);create_documents(run,sections,summary,tables,meta)
    if hashlib.sha256((root/config['arquivo']).read_bytes()).hexdigest()!=meta['sha256']:
        raise RuntimeError('A fonte foi alterada durante a execução. Rode novamente para obter um pacote consistente.')
    meta['fontes_referencia']=[]
    for pattern in ['manuais/*.pdf','relatorios base/**/*.pdf']:
        for source in root.glob(pattern):meta['fontes_referencia'].append({'arquivo':str(source.relative_to(root)),'sha256':hashlib.sha256(source.read_bytes()).hexdigest()})
    meta['criado_em']=datetime.now().astimezone().isoformat();meta['arquivos']=[{'caminho':str(p.relative_to(run)),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in sorted(run.rglob('*')) if p.is_file()]
    json_write(run/'manifesto.json',meta)
    (out/'ULTIMA_EXECUCAO.txt').write_text(run.name,encoding='utf-8')
    print(json.dumps({'pasta':str(run),'indicadores':summary},ensure_ascii=False,indent=2))

if __name__=='__main__':main()
