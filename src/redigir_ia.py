"""Redação opcional, separada dos cálculos e dos textos determinísticos."""
import argparse, json, urllib.request, re, hashlib
from pathlib import Path

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--config',default='config.exemplo.json'); args=ap.parse_args()
    root=Path(__file__).resolve().parents[1]
    config=json.loads((root/args.config).read_text(encoding='utf-8-sig'))
    out=root/config['saida']; run=out/(out/'ULTIMA_EXECUCAO.txt').read_text(encoding='utf-8').strip()
    ia=config['ia']
    if ia['provedor']!='ollama':
        raise SystemExit('Provedor não implementado. Use os prompts em docs/prompts para redação assistida em outro agente.')
    # O modelo não recebe dados reais. O corpo determinístico é inserido após a geração.
    prompt=(root/'docs/prompts/redacao_local.txt').read_text(encoding='utf-8')
    body={'model':ia['modelo'],'prompt':prompt,'stream':False,'think':False,'format':'json','options':{'temperature':0,'num_predict':500}}
    req=urllib.request.Request(ia['url'].rstrip('/')+'/api/generate',data=json.dumps(body).encode(),headers={'Content-Type':'application/json'})
    with urllib.request.urlopen(req,timeout=ia.get('timeout_segundos',300)) as resp: result=json.load(resp)
    generated=json.loads(result.get('response','{}'))
    if set(generated)!={'introducao','encerramento'}:raise SystemExit('Formato do Ollama inválido; nenhum texto foi aprovado.')
    for value in generated.values():
        if not isinstance(value,str) or re.search(r'\d|R\$|%',value) or len(value)>900:
            raise SystemExit('Ollama incluiu números ou formato inválido; nenhum texto foi aprovado.')
    original=(run/'textos/textos_para_RCI.txt').read_text(encoding='utf-8')
    text=('RASCUNHO PARA REVISÃO\nAbertura e encerramento sugeridos pelo Ollama; corpo e números gerados pelo motor determinístico.\n\n'+generated['introducao']+'\n\n'+original+'\n\n'+generated['encerramento'])
    (run/'textos/rascunho_ollama_REVISAR.txt').write_text(text,encoding='utf-8')
    (run/'dados/execucao_ia.json').write_text(json.dumps({'modelo':ia['modelo'],'provedor':'ollama','identificadores_enviados':False,'dados_reais_enviados':False,'numeros_no_texto_gerados_pelo_modelo':False,'eval_count':result.get('eval_count'),'total_duration_ns':result.get('total_duration')},indent=2),encoding='utf-8')
    manifest=json.loads((run/'manifesto.json').read_text(encoding='utf-8'))
    manifest['arquivos']=[{'caminho':str(p.relative_to(run)),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in sorted(run.rglob('*')) if p.is_file() and p.name!='manifesto.json']
    (run/'manifesto.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
    print('Rascunho local salvo em',run/'textos/rascunho_ollama_REVISAR.txt')

if __name__=='__main__': main()
