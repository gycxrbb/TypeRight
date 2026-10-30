$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class Win {
  public delegate bool CB(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(CB cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out R r);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out R r);
  [StructLayout(LayoutKind.Sequential)] public struct R { public int L, T, Rt, B; }
  public static IntPtr FindByTitlePart(uint pid, string part) {
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
  public static bool VisibleByTitlePart(uint pid, string part) {
    IntPtr h = FindByTitlePart(pid, part);
    return h != IntPtr.Zero && IsWindowVisible(h);
  }
}
"@
$shotDir = 'E:\TypeRight\tests\output'
function Capture($rect, $path) {
  $w = $rect.Rt - $rect.L; $h = $rect.B - $rect.T
  if ($w -le 0 -or $h -le 0) { throw "bad rect $($rect.L),$($rect.T) ${w}x${h}" }
  $bmp = New-Object System.Drawing.Bitmap $w, $h
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($rect.L, $rect.T, 0, 0, (New-Object System.Drawing.Size $w, $h))
  $g.Dispose(); $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
}
function CaptureWindow($hwnd, $path, $w, $h) {
  # PrintWindow 抓窗口内容（被遮挡也可用；PW_RENDERFULLCONTENT 支持 WPF/DirectComposition）
  if ($w -le 0 -or $h -le 0) { throw "bad window size ${w}x${h}" }
  $bmp = New-Object System.Drawing.Bitmap $w, $h
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $hdc = $g.GetHdc()
  [Win]::PrintWindow($hwnd, $hdc, 2) | Out-Null
  $g.ReleaseHdc($hdc); $g.Dispose()
  $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
}
function Analyze($path) {
  $bmp = New-Object System.Drawing.Bitmap $path
  $buckets = @{}
  for ($y = 0; $y -lt $bmp.Height; $y++) {
    for ($x = 0; $x -lt $bmp.Width; $x++) {
      $c = $bmp.GetPixel($x, $y)
      $k = (($c.R -shr 4) -shl 8) -bor (($c.G -shr 4) -shl 4) -bor ($c.B -shr 4)
      if ($buckets.ContainsKey($k)) { $buckets[$k]++ } else { $buckets[$k] = 1 }
    }
  }
  $bmp.Dispose()
  $top = $buckets.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
  $r = (($top.Key -shr 8) -band 0xF) * 17; $g = (($top.Key -shr 4) -band 0xF) * 17; $b = ($top.Key -band 0xF) * 17
  return @{ R = $r; G = $g; B = $b }
}
function CountNear($path, $tr, $tg, $tb, $tol) {
  $bmp = New-Object System.Drawing.Bitmap $path
  $n = 0
  for ($y = 0; $y -lt $bmp.Height; $y++) {
    for ($x = 0; $x -lt $bmp.Width; $x++) {
      $c = $bmp.GetPixel($x, $y)
      if ([math]::Abs($c.R - $tr) -le $tol -and [math]::Abs($c.G - $tg) -le $tol -and [math]::Abs($c.B - $tb) -le $tol) { $n++ }
    }
  }
  $bmp.Dispose(); return $n
}

# ---- 1. Word + doc ----
$src = Get-ChildItem 'E:\TypeRight\tests\*.docx' | Where-Object { $_.Name -notmatch '\u7edf\u4e00\u7248' } | Select-Object -First 1
Write-Output "DOC: $($src.Name)"
$word = New-Object -ComObject Word.Application
$word.Visible = $true
try { $word.WindowState = 1 } catch {}   # wdWindowStateMaximize
$doc = $word.Documents.Open($src.FullName, $false, $true)   # ReadOnly
Start-Sleep 3
try { $word.Activate() } catch {}
Start-Sleep 1

$addinObj = $null
try { $addinObj = $word.COMAddIns.Item('TypeRight.Connect').Object } catch {}
if ($null -eq $addinObj) { throw 'addin not loaded or Object null - check addin-log.txt' }
Write-Output 'ADDIN: loaded'

# ---- 2. trigger similar select via pending-action protocol ----
# 钩子挂起动作 → 用一次"选择变化"触发（WindowSelectionChange 在 Word 主 STA 线程执行 pending）
$pick = 0
for ($i = 3; $i -le 40; $i++) {
  $t = $doc.Paragraphs.Item($i).Range.Text
  if ($t.Length -ge 12) { $pick = $i; break }
}
if ($pick -eq 0) { throw 'no body paragraph found' }
Write-Output "SAMPLE paragraph: $pick"
$trigA = 2; $trigB = 4   # 交替用的触发段（与样本段不同）

function Trigger-Pending([int]$para) {
  $doc.Paragraphs.Item($para).Range.Select()
  Start-Sleep -Milliseconds 400
}

# 触发：先选样本段（设定样本），把 Word 顶到前台（悬浮条 Show+Activate 需要前台才能保持激活），再调用钩子
$doc.Paragraphs.Item($pick).Range.Select()
Start-Sleep -Milliseconds 600
try { $word.Activate() } catch {}
$wordHwnd = [IntPtr]$word.ActiveWindow.Hwnd
[Win]::ShowWindow($wordHwnd, 9) | Out-Null      # SW_RESTORE
[Win]::SetForegroundWindow($wordHwnd) | Out-Null
Start-Sleep -Milliseconds 600
$addinObj.AutomationSimilarSelect()

# ---- 3. capture floatbar（枚举 Word 进程窗口按标题匹配；首屏构造可达 ~10s）----
$wordPid = (Get-Process WINWORD | Select-Object -First 1).Id
$hwnd = [IntPtr]::Zero
for ($i = 0; $i -lt 40; $i++) {
  $hwnd = [Win]::FindByTitlePart($wordPid, 'TypeRight')
  if ($hwnd -ne [IntPtr]::Zero) { break }
  Start-Sleep -Milliseconds 500
}
if ($hwnd -eq [IntPtr]::Zero) { throw 'FloatBar window not found' }
$addinObj.AutomationShowFloatBar()
Start-Sleep -Milliseconds 1500
$rect = New-Object 'Win+R'
[Win]::GetWindowRect($hwnd, [ref]$rect) | Out-Null
$barVisible = [Win]::IsWindowVisible($hwnd)
Write-Output ("FLOATBAR rect: {0},{1} {2}x{3} vis={4}" -f $rect.L, $rect.T, ($rect.Rt - $rect.L), ($rect.B - $rect.T), $barVisible)

$results = @()
foreach ($case in @(
    @{ Name = 'dark';    Panel = @(42,42,42);   Text = @(242,242,242); Accent = @(77,142,247) },
    @{ Name = 'light';   Panel = @(251,251,252); Text = @(38,38,38);    Accent = @(43,87,154) },
    @{ Name = 'darkgray';Panel = @(74,74,74);   Text = @(245,245,245); Accent = @(122,165,236) }
  )) {
  # 本环境（COM 启动、后台）自动检测会走注册表兜底=浅色；三套主题全部显式强制后验证渲染
  $addinObj.AutomationForceFloatTheme($case.Name)
  Start-Sleep -Milliseconds 1500
  $path = Join-Path $shotDir "ui-float-$($case.Name).png"
  CaptureWindow $hwnd $path ($rect.Rt - $rect.L) ($rect.B - $rect.T)
  $dom = Analyze $path
  $panelOk = ([math]::Abs($dom.R - $case.Panel[0]) -le 12 -and [math]::Abs($dom.G - $case.Panel[1]) -le 12 -and [math]::Abs($dom.B - $case.Panel[2]) -le 12)
  # PrintWindow 对 WPF 分层窗口有青色透明环伪影：主导色被伪影劫持时，改按"面板色像素数"判定
  $panelPx = CountNear $path $case.Panel[0] $case.Panel[1] $case.Panel[2] 8
  if (-not $panelOk -and $panelPx -ge 30000) { $panelOk = $true }
  $textPx = CountNear $path $case.Text[0] $case.Text[1] $case.Text[2] 30
  $accPx = CountNear $path $case.Accent[0] $case.Accent[1] $case.Accent[2] 40
  $verdict = if ($panelOk -and $textPx -ge 60 -and $accPx -ge 15) { 'PASS' } else { 'FAIL' }
  $results += $verdict
  Write-Output ("{0,-9} panel=({1},{2},{3}) panelPx={4} textPx={5} accentPx={6} -> {7}" -f $case.Name, $dom.R, $dom.G, $dom.B, $panelPx, $textPx, $accPx, $verdict)
}

# word window context shot (ribbon + doc + floatbar, PrintWindow 不受遮挡影响)
$wh = [IntPtr]$word.ActiveWindow.Hwnd
$wr = New-Object 'Win+R'
[Win]::GetWindowRect($wh, [ref]$wr) | Out-Null
CaptureWindow $wh (Join-Path $shotDir 'ui-word-context.png') ($wr.Rt - $wr.L) ([math]::Min($wr.B - $wr.T, 760))

# ---- 4. cleanup ----
$addinObj.AutomationCloseSession()
Start-Sleep 1
$doc.Close(0)
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output ('SUMMARY: ' + (($results | Where-Object { $_ -eq 'FAIL' }).Count) + ' fail / ' + $results.Count)
