# 用用户的真实文档复现两个 bug：9s 耗时分解 + 图标消失取证
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class RB {
  public delegate bool CB(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(CB cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out R r);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [StructLayout(LayoutKind.Sequential)] public struct R { public int L, T, Rt, B; }
  public static IntPtr MainOf(uint pid) {
    IntPtr found = IntPtr.Zero;
    EnumWindows(delegate(IntPtr h, IntPtr l) {
      uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid) { var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
        if (sb.ToString().IndexOf("Word", StringComparison.OrdinalIgnoreCase) >= 0) { found = h; return false; } }
      return true;
    }, IntPtr.Zero);
    return found;
  }
}
"@
function ShotRibbon($hwnd, $path) {
  $r = New-Object RB+R
  [RB]::GetWindowRect($hwnd, [ref]$r) | Out-Null
  $w = $r.Rt - $r.L; $h = [Math]::Min(220, $r.B - $r.T)
  $bmp = New-Object System.Drawing.Bitmap($w, $h)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $hdc = $g.GetHdc(); [RB]::PrintWindow($hwnd, $hdc, 2) | Out-Null; $g.ReleaseHdc($hdc); $g.Dispose()
  $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
}

$word = New-Object -ComObject Word.Application
$word.Visible = $true
Start-Sleep -Seconds 7
$doc = $word.Documents.Open('E:\TypeRight\tests\第二十六章 二次函数-学生.docx')
Start-Sleep -Seconds 4
# 放大 + 前台（保证检测/截图稳定）
$proc = Get-Process WINWORD | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
[RB]::ShowWindow($proc.MainWindowHandle, 3) | Out-Null
[RB]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
Start-Sleep -Seconds 2
ShotRibbon $proc.MainWindowHandle 'E:\TypeRight\spikes\ribbon-before.png'
Write-Host "已截选前功能区"

# 光标放到正文段（找含"解"或数字的长段）
$paras = $doc.Paragraphs
$p = $null
for ($i = 10; $i -le 80; $i++) {
  $pp = $paras.Item($i)
  $t = [string]$pp.Range.Text
  if ($t.Length -gt 20 -and $t -match '解|抛物线|y\s*=') { $p = $pp; Write-Host "样本段 #$i : $($t.Substring(0, [Math]::Min(20, $t.Length)))"; break }
}
$word.Selection.SetRange($p.Range.Start, $p.Range.Start)
Start-Sleep -Seconds 1
# 计时同类选择（端到端）
$t0 = Get-Date
$word.COMAddIns.Item('TypeRight.Connect').Object.AutomationSimilarSelect()
Start-Sleep -Seconds 1
# 等高亮+浮条稳定（轮询浮条窗口出现）
$deadline = (Get-Date).AddSeconds(30)
while ((Get-Date) -lt $deadline) {
  $fb = $null
  foreach ($w in (Get-Process WINWORD | Where-Object { $_.Id -eq $proc.Id })) {}
  Start-Sleep -Milliseconds 300
  break
}
Start-Sleep -Seconds 2
$t1 = Get-Date
Write-Host ("端到端（含 2s 余量）: {0:N1}s" -f ($t1 - $t0).TotalSeconds)
ShotRibbon $proc.MainWindowHandle 'E:\TypeRight\spikes\ribbon-after.png'
Write-Host "已截选后功能区"
$doc.Close(0)
$word.Quit()
Write-Host "── 日志新增 ──"
$n = (Get-Content "$env:LOCALAPPDATA\TypeRight\addin-log.txt" | Measure-Object -Line).Lines
# 简单取最后 40 行
Get-Content "$env:LOCALAPPDATA\TypeRight\addin-log.txt" -Tail 40
