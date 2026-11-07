# 图标消失复现：先开 TypeRight 选项卡（图标渲染）→ 同类选择 → 截图对比 + GetImage 日志
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class RB2 {
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out R r);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [StructLayout(LayoutKind.Sequential)] public struct R { public int L, T, Rt, B; }
}
"@
function ShotTop($hwnd, $path) {
  $r = New-Object RB2+R
  [RB2]::GetWindowRect($hwnd, [ref]$r) | Out-Null
  $w = $r.Rt - $r.L; $h = [Math]::Min(260, $r.B - $r.T)
  $bmp = New-Object System.Drawing.Bitmap($w, $h)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $hdc = $g.GetHdc(); [RB2]::PrintWindow($hwnd, $hdc, 2) | Out-Null; $g.ReleaseHdc($hdc); $g.Dispose()
  $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
}
$word = New-Object -ComObject Word.Application
$word.Visible = $true
Start-Sleep -Seconds 7
$doc = $word.Documents.Open('E:\TypeRight\tests\第二十六章 二次函数-学生.docx', $null, $true)
Start-Sleep -Seconds 3
$proc = Get-Process WINWORD | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
[RB2]::ShowWindow($proc.MainWindowHandle, 3) | Out-Null
[RB2]::SetForegroundWindow($proc.MainWindowHandle) | Out-Null
Start-Sleep -Seconds 2

# 打开 TypeRight 选项卡（触发图标懒加载渲染）
$root = [System.Windows.Automation.AutomationElement]::RootElement
$win = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $proc.Id)))
$tabs = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::TabItem)))
for ($i = 0; $i -lt $tabs.Count; $i++) {
  if ($tabs[$i].Current.Name -eq 'TypeRight') {
    $sel = $tabs[$i].GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    $sel.Select()
    Write-Host "TypeRight 选项卡已打开"
    break
  }
}
Start-Sleep -Seconds 3
ShotTop $proc.MainWindowHandle 'E:\TypeRight\spikes\icons-before.png'

# 同类选择（触发 RefreshRibbonTheme → 可能 Invalidate）
$paras = $doc.Paragraphs
$p = $null
for ($i = 10; $i -le 80; $i++) {
  $pp = $paras.Item($i)
  $t = [string]$pp.Range.Text
  if ($t.Length -gt 20 -and $t -match '解|抛物线|y\s*=') { $p = $pp; break }
}
$word.Selection.SetRange($p.Range.Start, $p.Range.Start)
Start-Sleep -Seconds 1
$word.COMAddIns.Item('TypeRight.Connect').Object.AutomationSimilarSelect()
Start-Sleep -Seconds 6
ShotTop $proc.MainWindowHandle 'E:\TypeRight\spikes\icons-after.png'
Write-Host "前后对比图已存"
$doc.Close(0)
$word.Quit()
Write-Host "── GetImage 日志 ──"
Get-Content "$env:LOCALAPPDATA\TypeRight\addin-log.txt" -Tail 25 | Select-String -Pattern 'GetImage|RefreshRibbon|Invalidate|SimilarSelect: enter|匹配'
