param([string]$Config = 'config.local.json', [switch]$TextoIA)
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
if (-not (Test-Path -LiteralPath $Config)) { $Config = 'config.exemplo.json' }
$taskPython = if ($env:ADIANTAMENTO_PYTHON) { $env:ADIANTAMENTO_PYTHON } elseif (Test-Path -LiteralPath '.venv/Scripts/python.exe') { Join-Path $PSScriptRoot '.venv/Scripts/python.exe' } else { 'python' }
& $taskPython src/analisador.py --config $Config
if ($LASTEXITCODE -ne 0) { throw 'Falha na análise. Consulte a mensagem acima.' }
& $taskPython src/aprofundar_relatorio.py --config $Config --mesma-execucao
if ($LASTEXITCODE -ne 0) { throw 'Quadros gerados; houve falha na análise textual aprofundada.' }
& $taskPython src/complementos.py --config $Config
if ($LASTEXITCODE -ne 0) { throw 'Houve falha nos quadros mensais ou na relação individual.' }
if ($TextoIA) {
  & $taskPython src/redigir_ia.py --config $Config
  if ($LASTEXITCODE -ne 0) { throw 'Análise gerada; a redação por IA falhou.' }
}
