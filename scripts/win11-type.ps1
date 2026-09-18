# Type into the running corro GUI and screenshot the result.
param([string]$Shot = "C:\Windows\Temp\corro-gui-type.png")
$ws = New-Object -ComObject WScript.Shell
$ok = $ws.AppActivate("corro 0.7.0")
"activated=$ok"
Start-Sleep -Seconds 1
$ws.SendKeys("hi")
Start-Sleep -Milliseconds 500
$ws.SendKeys("{ENTER}")
Start-Sleep -Seconds 2
& "$PSScriptRoot\win11-shot.ps1" -Out $Shot
