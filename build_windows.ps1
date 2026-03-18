$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

if (-not (Get-Command py -ErrorAction SilentlyContinue)) {
  throw "Python launcher 'py' no encontrado. Instala Python 3 para Windows con PATH/launcher."
}

py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt pyinstaller

.\.venv\Scripts\pyinstaller.exe --noconsole --onefile --name DivoomKeeper app.py

Write-Host "Build completado: $PSScriptRoot\dist\DivoomKeeper.exe" -ForegroundColor Green
