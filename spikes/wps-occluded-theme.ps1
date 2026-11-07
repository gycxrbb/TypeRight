# 遮挡场景产品级主题验证：WPS（亮）被控制台遮挡 → SimilarSelect → PrintWindow 抓浮条
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class PW3 {
  public delegate bool CB(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(CB cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out R r);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
  [StructLayout(LayoutKind.Sequential)] public struct R { public int L, T, Rt, B; }
  public static IntPtr FindByTitle(uint pid, string part) {
    IntPtr found = IntPtr.Zero;
    EnumWindows(delegate(IntPtr h, IntPtr l) {
      uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid) {
        var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
        if (sb.ToString().IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0) { found = h; return false; }
      }
      return true;
    }, IntPtr.Zero);
    return found;
  }
}
"@
$wps = New-Object -ComObject 'KWPS.Application'
$wps.Visible = $true
Start-Sleep -Seconds 8
$doc = $wps.Documents.Open("$env:TEMP\wps-smoke.docx")
Start-Sleep -Seconds 3
$paras = $doc.Paragraphs
$p = $null
for ($i = 20; $i -le 60; $i++) {
  $pp = $paras.Item($i)
  $t = [string]$pp.Range.Text
  if ($t.Length -gt 20 -and $t -match '方程|函数|解') { $p = $pp; break }
}
$wps.Selection.SetRange($p.Range.Start, $p.Range.Start)
Start-Sleep -Seconds 1
$wps.COMAddIns.Item('TypeRight.Connect').Object.AutomationSimilarSelect()
Start-Sleep -Seconds 7
$proc = Get-Process wps | Sort-Object StartTime -Descending | Select-Object -First 1
$h = [PW3]::FindByTitle($proc.Id, 'TypeRight')
Write-Host "浮条 hwnd: $h"
if ($h -ne [IntPtr]::Zero) {
  $r = New-Object PW3+R
  [PW3]::GetWindowRect($h, [ref]$r) | Out-Null
  $w = $r.Rt - $r.L; $hh = $r.B - $r.T
  $bmp = New-Object System.Drawing.Bitmap($w, $hh)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $hdc = $g.GetHdc()
  [PW3]::PrintWindow($h, $hdc, 2) | Out-Null
  $g.ReleaseHdc($hdc); $g.Dispose()
  $out = 'E:\TypeRight\spikes\wps-occluded-float.png'
  $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
  Write-Host "已抓: $out (${w}x${hh})（WPS 此时被控制台遮挡）"
}
$doc.Close(0)
try { $wps.Quit() } catch {}
