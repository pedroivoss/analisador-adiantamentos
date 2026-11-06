# Execução no computador atual

A instalação de Python não foi encontrada no PATH durante a preparação. O Codex fornece um runtime em sua instalação empacotada. O nome do diretório pode mudar em atualizações; consulte novamente o runtime se deixar de existir.

```powershell
$taskRuntime = 'C:/Program Files/WindowsApps/OpenAI.CodexPrimaryRuntime.v26-1104-2876-0_26.1104.2876.0_x64__3k8sg7r9htsxt/dependencies'
$env:ADIANTAMENTO_PYTHON = "$taskRuntime/python/python.exe"
$env:ADIANTAMENTO_NODE = "$taskRuntime/node/bin/node.exe"
Set-Location D:/Projetos/analisador-adiantamento
.\analisar.ps1
# Para redação local adicional:
& $env:ADIANTAMENTO_PYTHON src/redigir_ia.py --config config.local.json
```

O matplotlib foi instalado em `.runtime-deps`, sem modificar as bibliotecas gerenciadas do Codex. O motor encontra essa pasta automaticamente. A instalação é ignorada pelo Git. Na migração, recrie um `.venv` com `requirements.txt`. O `.runtime-deps` pode conter binários específicos da versão de Python e do Windows; não deve ser copiado para outras plataformas.

Na execução assistida, a planilha é exportada pelo artifact-tool do Codex. `node_modules` é uma junção local para o runtime e fica fora do versionamento. Em um ambiente Python independente, sem `ADIANTAMENTO_NODE`, o motor usa openpyxl somente para criar a planilha analítica nova. A fonte nunca é salva. Não configure `ADIANTAMENTO_NODE` em outra máquina sem garantir a disponibilidade de `@oai/artifact-tool`.

O aviso do openpyxl sobre extensão de validação de dados ocorre na leitura de recursos específicos do Excel. O código não salva a fonte; portanto não remove as validações do arquivo original.

O modo IA consulta `/api/generate` no Ollama e grava o rascunho na última execução concluída. Usa modelo informado em `config.local.json`, temperatura zero, JSON e limite de geração. O modelo gera só abertura/encerramento sem números e não recebe dados reais; o corpo é inserido pelo código. Confirme primeiro que `http://localhost:11434/api/tags` lista o modelo. Não há download automático de modelos.
