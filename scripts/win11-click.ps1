# Click at screen coords, then screenshot.
param([int]$X = 500, [int]$Y = 500, [string]$Shot = "C:\Windows\Temp\corro-click.png")
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Mouse {
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
  [DllImport("user32.dll")] public static extern void mouse_event(int flags, int dx, int dy, int data, IntPtr extra);
}
"@
[Mouse]::SetCursorPos($X, $Y)
Start-Sleep -Milliseconds 300
[Mouse]::mouse_event(0x0002, 0, 0, 0, [IntPtr]::Zero)
Start-Sleep -Milliseconds 100
[Mouse]::mouse_event(0x0004, 0, 0, 0, [IntPtr]::Zero)
Start-Sleep -Milliseconds 500
& "$PSScriptRoot\win11-shot.ps1" -Out $Shot
