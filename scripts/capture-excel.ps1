param([int]$ExcelProcessId, [string]$OutputPath)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;
using System.Runtime.InteropServices;
public class ExcelCapture {
 [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
 [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
 public struct Rect {public int Left,Top,Right,Bottom;}
}
'@
$process = Get-Process -Id $ExcelProcessId
[void][ExcelCapture]::SetThreadDpiAwarenessContext([IntPtr](-4))
if ($process.ProcessName -ne 'EXCEL') { throw 'Target process must be Excel.' }
$windowRect=New-Object ExcelCapture+Rect
if(-not [ExcelCapture]::GetWindowRect($process.MainWindowHandle,[ref]$windowRect)){throw 'Excel window was not found.'}
$bitmap=New-Object Drawing.Bitmap ([int]($windowRect.Right-$windowRect.Left)),([int]($windowRect.Bottom-$windowRect.Top))
$graphics=[Drawing.Graphics]::FromImage($bitmap)
try {
 $hdc=$graphics.GetHdc()
 try {if(-not [ExcelCapture]::PrintWindow($process.MainWindowHandle,$hdc,2)){throw 'Window capture failed.'}}finally{$graphics.ReleaseHdc($hdc)}
 $bitmap.Save($OutputPath)
}finally{$graphics.Dispose();$bitmap.Dispose()}
Write-Output ('Captured Excel window: ' + $OutputPath)
