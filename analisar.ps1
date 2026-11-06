param([string]$Config = 'config.local.json', [switch]$TextoIA)
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
if (-not (Test-Path -LiteralPath $Config)) { $Config = 'config.exemplo.json' }
$taskPython = if ($env:ADIANTAMENTO_PYTHON) { $env:ADIANTAMENTO_PYTHON } elseif (Test-Path -LiteralPath '.venv/Scripts/python.exe') { Join-Path $PSScriptRoot '.venv/Scripts/python.exe' } else { 'python' }
& $taskPython src/analisador.py --config $Config
if ($LASTEXITCODE -ne 0) { throw 'Falha na análise. Consulte a mensagem acima.' }
if ($TextoIA) {
  & $taskPython src/redigir_ia.py --config $Config
  if ($LASTEXITCODE -ne 0) { throw 'Análise gerada; a redação por IA falhou.' }
}
