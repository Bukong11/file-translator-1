$ErrorActionPreference = "Stop"

$appRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$envFile = Join-Path $appRoot ".openai-env.bat"
if (-not (Test-Path $envFile)) {
  Write-Host "Missing .openai-env.bat"
  Read-Host "Press Enter to exit"
  exit 1
}

foreach ($line in Get-Content $envFile) {
  if ($line -match '^set\s+"?OPENAI_API_KEY=(.+?)"?\s*$') {
    $env:OPENAI_API_KEY = $Matches[1].Trim('"')
  }
}

$env:OPENAI_MODEL = "gpt-5.5"
Set-Location (Join-Path $appRoot "backend")
.\.venv\Scripts\python.exe -m uvicorn main:app --host 0.0.0.0 --port 8001
