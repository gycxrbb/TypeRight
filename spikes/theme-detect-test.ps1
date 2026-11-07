# 复现 OfficeTheme.DetectByPixels：对 Word(暗色) 与 WPS(亮色) 实测采样结果
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class OW {
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out R r);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
  [StructLayout(LayoutKind.Sequential)] public struct R { public int L, T, Rt, B; }
}
"@
function Sample-Host($procName, $label) {
  $p = Get-Process $procName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
  if (-not $p) { Write-Host "$label : 无可见窗口"; return }
  $r = New-Object OW+R
  [OW]::GetWindowRect($p.MainWindowHandle, [ref]$r) | Out-Null
  $w = $r.Rt - $r.L; $h = $r.B - $r.T
  Write-Host "$label : ${w}x${h} visible=$([OW]::IsWindowVisible($p.MainWindowHandle)) iconic=$([OW]::IsIconic($p.MainWindowHandle))"
  $bmp = New-Object System.Drawing.Bitmap(1,1)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  foreach ($frac in 0.03, 0.07, 0.13) {
    $y = $r.T + [int]($h * $frac)
    $vals = @()
    for ($fx = 0; $fx -lt 12; $fx++) {
      $x = $r.L + [int]($w * 0.15) + [int]($w * 0.70) * $fx / 11
      $g.CopyFromScreen($x, $y, 0, 0, (New-Object System.Drawing.Size(1,1)))
      $vals += $bmp.GetPixel(0,0)
    }
    $avg = @{ R = ($vals | Measure-Object R -Average).Average; G = ($vals | Measure-Object G -Average).Average; B = ($vals | Measure-Object B -Average).Average }
    $lum = ($avg.R * 0.299 + $avg.G * 0.587 + $avg.B * 0.114) / 255
    $kind = if ($lum -lt 0.45) { 'Dark' } elseif ($lum -lt 0.75) { 'DarkGray' } else { 'Light' }
    Write-Host ("  行 {0:p0}: 均色 #{1:X2}{2:X2}{3:X2} 亮度={4:N2} → {5}" -f $frac, [int]$avg.R, [int]$avg.G, [int]$avg.B, $lum, $kind)
  }
  $g.Dispose(); $bmp.Dispose()
}
# 拉 Word（本机暗色主题）
$word = New-Object -ComObject Word.Application
$word.Visible = $true
Start-Sleep -Seconds 6
$null = $word.Documents.Add()
Start-Sleep -Seconds 3
Sample-Host 'WINWORD' 'Word(期望 Dark)'
$word.Quit()
# 拉 WPS（本机亮色）
$wps = New-Object -ComObject 'KWPS.Application'
$wps.Visible = $true
Start-Sleep -Seconds 8
$null = $wps.Documents.Add()
Start-Sleep -Seconds 3
Add-Type @"
using System.Runtime.InteropServices;
public static class FW2 { [DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h); [DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr h, int c); }
"@
$pw = Get-Process wps | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
[FW2]::ShowWindow($pw.MainWindowHandle, 3) | Out-Null
[FW2]::SetForegroundWindow($pw.MainWindowHandle) | Out-Null
Start-Sleep -Seconds 3
Sample-Host 'wps' 'WPS前台最大化(期望 Light)'
try { $wps.Quit() } catch {}
