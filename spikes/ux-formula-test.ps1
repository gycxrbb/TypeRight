# 公式转换体验验证：① 正常转换（进度窗出现/过程态截图/254 处完成） ② Esc 取消（部分完成+不崩溃）
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class PWX {
  public delegate bool CB(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(CB cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out R r);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
  [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
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
  public static void Shot(IntPtr h, string path) {
    var r = new R(); GetWindowRect(h, out r);
    int w = r.Rt - r.L, hh = r.B - r.T;
    if (w -le 0 -or hh -le 0) { return }
    var bmp = new System.Drawing.Bitmap(w, hh);
    var g = System.Drawing.Graphics.FromImage(bmp);
    var hdc = g.GetHdc(); PrintWindow(h, hdc, 2); g.ReleaseHdc(hdc); g.Dispose();
    bmp.Save(path, System.Drawing.Imaging.ImageFormat::Png); bmp.Dispose();
  }
}
"@

# ── 场景 1：正常转换（讲义副本，254 处） ──
$word = New-Object -ComObject Word.Application
$word.Visible = $true
Start-Sleep -Seconds 7
$copy1 = "$env:TEMP\ux-full.docx"
Copy-Item 'E:\TypeRight\tests\初三数学代数终极讲义_一元二次方程与二次函数大一统.docx' $copy1 -Force
Set-ItemProperty $copy1 -Name IsReadOnly -Value $false
$doc = $word.Documents.Open($copy1)
Start-Sleep -Seconds 3
$proc = Get-Process WINWORD | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1

$word.COMAddIns.Item('TypeRight.Connect').Object.AutomationConvertFormulas()
# 轮询抓进度窗（过程态）
$found = $false
for ($i = 0; $i -lt 30; $i++) {
  Start-Sleep -Milliseconds 200
  $h = [PWX]::FindByTitle($proc.Id, 'TypeRightProgress')
  if ($h -ne [IntPtr]::Zero) { [PWX]::Shot($h, 'E:\TypeRight\spikes\ux-progress-mid.png'); $found = $true; break }
}
Write-Host "进度窗出现: $found"
Start-Sleep -Seconds 8
$h2 = [PWX]::FindByTitle($proc.Id, 'TypeRightProgress')
if ($h2 -ne [IntPtr]::Zero) { [PWX]::Shot($h2, 'E:\TypeRight\spikes\ux-progress-done.png'); Write-Host "完成态已截" }
$z1 = $doc.OMaths.Count
Write-Host "场景1 转换后公式数: $z1 (期望约 254)"
$doc.Close(0)

# ── 场景 2：Esc 取消（部分完成 + 无崩溃） ──
$copy2 = "$env:TEMP\ux-cancel.docx"
Copy-Item 'E:\TypeRight\tests\初三数学代数终极讲义_一元二次方程与二次函数大一统.docx' $copy2 -Force
Set-ItemProperty $copy2 -Name IsReadOnly -Value $false
$doc2 = $word.Documents.Open($copy2)
Start-Sleep -Seconds 3
$word.COMAddIns.Item('TypeRight.Connect').Object.AutomationConvertFormulas()
Start-Sleep -Milliseconds 1400
# 按住 Esc 800ms（GetAsyncKeyState 读实时按下态，检查点每 16 处一次必命中）
[PWX]::keybd_event(0x1B, 0, 0, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 800
[PWX]::keybd_event(0x1B, 0, 2, [UIntPtr]::Zero)
Start-Sleep -Seconds 6
$z2 = $doc2.OMaths.Count
Write-Host "场景2 取消后公式数: $z2 (期望介于 0 与 254 之间的部分值)"
Start-Sleep -Seconds 2
$h3 = [PWX]::FindByTitle($proc.Id, 'TypeRightProgress')
if ($h3 -ne [IntPtr]::Zero) { [PWX]::Shot($h3, 'E:\TypeRight\spikes\ux-progress-cancel.png'); Write-Host "取消态已截" }
$doc2.Close(0)
$word.Quit()
Write-Host "── 日志尾部 ──"
Get-Content "$env:LOCALAPPDATA\TypeRight\addin-log.txt" -Tail 6
