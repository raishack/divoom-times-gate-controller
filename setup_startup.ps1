$ErrorActionPreference = 'Stop'
$exe = 'C:\Users\raish\Desktop\DivoomKeeper\dist\DivoomKeeper.exe'
if (-not (Test-Path $exe)) { throw "No existe $exe" }

$startup = [Environment]::GetFolderPath('Startup')
$link = Join-Path $startup 'DivoomKeeper.lnk'

$w = New-Object -ComObject WScript.Shell
$s = $w.CreateShortcut($link)
$s.TargetPath = $exe
$s.WorkingDirectory = Split-Path $exe
$s.WindowStyle = 7
$s.IconLocation = $exe
$s.Description = 'Divoom Keeper tray auto-start'
$s.Save()

Write-Output "Shortcut: $link"
Write-Output "Target: $exe"

Start-Process -FilePath $exe
Write-Output 'DivoomKeeper launched'
