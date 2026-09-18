# Launch gcorro --gui with stderr to a log file (for TMP traces).
param([string]$Exe = "C:\corro-test\gcorro.exe", [string]$Err = "C:\corro-test\err.log")
Stop-Process -Name gcorro -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
if (Test-Path $Err) { Remove-Item $Err -Force }
$proc = Start-Process -FilePath $Exe -ArgumentList "--gui" -RedirectStandardError $Err -PassThru
Start-Sleep -Seconds 12
& "$PSScriptRoot\win11-shot.ps1" -Out "C:\Windows\Temp\corro-rel.png"
$alive = -not $proc.HasExited
"alive=$alive pid=$($proc.Id)"
