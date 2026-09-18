# Launch gcorro --gui on Win11, screenshot, report.
# Usage: powershell -ExecutionPolicy Bypass -File rungui.ps1 -Exe <path> -Shot <png> [-Stay]
param([string]$Exe = "C:\corro-test\gcorro.exe", [string]$Shot = "C:\Windows\Temp\corro-gui.png", [switch]$Stay)
New-Item -ItemType Directory -Force -Path (Split-Path $Exe) | Out-Null
Stop-Process -Name gcorro -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
$proc = Start-Process -FilePath $Exe -ArgumentList "--gui" -PassThru
Start-Sleep -Seconds 12
& "$PSScriptRoot\win11-shot.ps1" -Out $Shot
$alive = -not $proc.HasExited
"alive=$alive pid=$($proc.Id)"
if (-not $Stay) {
  if (-not $proc.HasExited) { Stop-Process -Id $proc.Id -Force }
  "closed"
}
