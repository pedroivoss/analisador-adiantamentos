# Analisador de adiantamentos de Ubatuba

Gera material editável para a seção de adiantamentos do Relatório de Controle Interno (RCI), a partir da planilha de prestação de contas. A fonte original é somente lida. O projeto começa com um teste do primeiro semestre de 2026 e pode ser configurado para outros bimestres, semestres e exercícios.

## Começar

Requer Python 3.11 ou superior. No PowerShell, dentro desta pasta:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item config.exemplo.json config.local.json
.\analisar.ps1
```

Não repita o `Copy-Item` se já personalizou a configuração. Se a política de execução impedir o script, rode diretamente:

```powershell
.venv\Scripts\python.exe src/analisador.py --config config.local.json
```

Nesta sessão do Codex, os executáveis empacotados foram usados; `docs/EXECUCAO_LOCAL.md` registra como repetir sem instalar outro Python. Para migrar, prefira o ambiente virtual acima.

## Entregas

Cada execução cria uma pasta nova em `Analises/2026/Jan-jun/relatorios gerados/`. O arquivo `ULTIMA_EXECUCAO.txt` aponta para a última execução concluída. Não é necessário apagar relatórios: alterações manuais anteriores ficam preservadas.

- Guia geral em Word e PDF, explicando pastas, arquivos e uso.
- Relatório com seis partes em Word e PDF.
- Cinco gráficos PNG e SVG e textos por assunto.
- Planilha analítica editável e tabelas CSV, incluindo rankings completos, médias, mediana, extremos, prazos e exames cruzados.
- Base normalizada, alertas por linha, conciliação e manifesto com hashes.
- Redação opcional pelo Ollama, separada do relatório determinístico.

## IA local

Os cálculos e textos de acompanhamento funcionam sem modelo de linguagem. Para gerar também uma proposta de redação local:

```powershell
.\analisar.ps1 -TextoIA
# Ou apenas redigir a execução já concluída:
.venv\Scripts\python.exe src/redigir_ia.py --config config.local.json
```

O padrão é Ollama em `http://localhost:11434`, modelo `qwen3:14b`. `qwen3:8b` também foi encontrado instalado. Troque `ia.modelo` na configuração para comparar. Não há limite de solicitações cobrado por API nesse fluxo local; tempo e capacidade dependem da máquina. Nesta versão a IA sugere apenas abertura e encerramento, sem receber dados reais nem redigir números. O corpo é inserido pelo motor determinístico. Essa restrição foi adotada porque o primeiro teste livre produziu erros numéricos. A minuta precisa de conferência. Nenhuma conta externa ou browser é acessado automaticamente.

Para Codex, Claude Code ou Gemini, use o prompt de `docs/prompts/continuar_projeto.txt`. Integrações automáticas com APIs externas ainda não estão implementadas; não basta trocar o nome do provedor na configuração.

## Revisão antes do RCI

A contagem principal é de registros por nota/natureza; processos SEI distintos aparecem separadamente. Não é possível deduzir a quantidade exata de solicitações só com essas colunas. As durações de aplicação usam a **última aplicação**, pois a primeira despesa não está disponível. Os exames SMF e CGM representam a prestação de contas.

Leia `docs/METODOLOGIA.md`, confira `dados/alertas.csv` e valide o calendário municipal. Os atrasos divulgados se referem aos vencimentos informados na planilha. A comparação técnica de dias úteis não substitui esses vencimentos e não é um julgamento automático.

## Versionamento

O `.gitignore` exclui as fontes reais, manuais, relatórios base, relatórios gerados, configuração local e ambientes. Código, documentação, configuração de exemplo e testes podem ser versionados. O Git não copia dados ignorados entre computadores: mova as fontes separadamente, mantendo as pastas relativas. Não há upload nem repositório remoto configurado por este projeto.

```powershell
python -m unittest discover -s tests -v
```
