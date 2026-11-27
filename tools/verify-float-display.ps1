# 浮条字体/字号显示取证（真机 COM，PowerShell 后期绑定）。
# 断言的是"浮条上真的显示了什么"，不是内部字段——字段对了但没渲染出来也是 bug。
# 前置：加载项已部署且未运行 Word（本脚本自己起关 Word，不动别人的窗口）。
$ErrorActionPreference = 'Stop'
[int]$fails = 0
$word = $null
$doc = $null

function Check([string]$actual, [string]$mustContain, [string]$what) {
  if ($actual -like "*$mustContain*") { Write-Host "  [PASS] $what"; return 0 }
  Write-Host "  [FAIL] $what | 期望含「$mustContain」实际「$actual」"; return 1
}
function CheckNot([string]$actual, [string]$mustNot, [string]$what) {
  if ($actual -notlike "*$mustNot*") { Write-Host "  [PASS] $what"; return 0 }
  Write-Host "  [FAIL] $what | 不应含「$mustNot」实际「$actual」"; return 1
}

# 先写全文再整段刷格式：逐段取 Range 会拿到"刚插入后的下一段"，把格式写到空段上
function BuildSample([string]$font, [double]$size, [int]$n, [bool]$pattern) {
  $script:doc.Content.Delete()
  $sb = New-Object System.Text.StringBuilder
  for ($i = 1; $i -le $n; $i++) {
    $text = if ($pattern) { "第${i}题：下列说法正确的是（　）" } else { "这是第 $i 行同类段落，用于浮条字体字号显示取证。" }
    [void]$sb.Append($text + "`r")
  }
  $script:doc.Content.InsertAfter($sb.ToString())
  $r = $script:doc.Content
  $r.Font.NameFarEast = $font
  $r.Font.NameAscii = $font
  $r.Font.Size = [single]$size
}

function SnapAt([string]$label) {
  # 首次构造悬浮条要起 WPF 窗口（实测可达数秒），固定等待会读到"刚 Reset 完"的空态。
  # 轮询到字号有值为止（最多 12s），再取稳定快照。
  $s = ''
  for ($i = 0; $i -lt 24; $i++) {
    Start-Sleep -Milliseconds 500
    $raw = @($script:addin.AutomationFloatBarDisplay())
    $cand = [string]($raw | Where-Object { "$_" -match '\|' } | Select-Object -Last 1)
    if ($cand) { $s = $cand }
    if ($s -and $s -notmatch '\|—$' -and $s -notmatch '^\|') { break }
  }
  Write-Host "  显示快照($label): $s"
  return $s
}

try {
  Write-Host "── 浮条字体/字号显示取证 ──────────────────────"
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0

  $addin = $null
  try { $addin = $word.COMAddIns.Item('TypeRight.Connect').Object } catch { }
  if ($null -eq $addin) { Write-Host "  [FATAL] 加载项未连接（检查 addin-log.txt）"; exit 2 }
  Write-Host "  ADDIN: loaded"
  $script:addin = $addin
}
catch {
  Write-Host "  [FATAL] $($_.Exception.Message)"
  exit 2
}

try {
  $doc = $word.Documents.Add()
  $script:doc = $doc

  # 预热：加载项刚连接时第一次 Automation 调用会被 SetTimer 时序吃掉（实测 A 用例静默失败）。
  # 先空跑一次同类选择把绳拉直，进第一个真用例前关掉会话。
  BuildSample '仿宋' 14 2 $false
  $word.Selection.SetRange(1, 1)
  $addin.AutomationSimilarSelect() | Out-Null
  Start-Sleep -Milliseconds 2500
  $addin.AutomationCloseSession() | Out-Null
  Start-Sleep -Milliseconds 700

  # ── A：仿宋 14pt → 「四号」+ 字体回填（复现用户截图场景）──
  BuildSample '仿宋' 14 3 $false
  $word.Selection.SetRange(1, 1)
  $addin.AutomationSimilarSelect() | Out-Null
  $snap = SnapAt '仿宋14pt'
  $fails += Check $snap '四号' 'A 14pt 应显示中文「四号」'
  $fails += Check $snap '仿宋' 'A 字体框应回填样本字体「仿宋」'
  $fails += CheckNot $snap '|0' 'A 字号不得显示 0'
  $fails += CheckNot $snap '|—' 'A 字号不得是未知占位'
  $addin.AutomationCloseSession() | Out-Null
  Start-Sleep -Milliseconds 400

  # ── B：宋体 12pt → 「小四」──
  BuildSample '宋体' 12 3 $false
  $word.Selection.SetRange(1, 1)
  $addin.AutomationSimilarSelect() | Out-Null
  $snap = SnapAt '宋体12pt'
  $fails += Check $snap '小四' 'B 12pt 应显示中文「小四」'
  $addin.AutomationCloseSession() | Out-Null
  Start-Sleep -Milliseconds 400

  # ── C：宋体 10.5pt → 「五号」（最常见正文字号）──
  BuildSample '宋体' 10.5 3 $false
  $word.Selection.SetRange(1, 1)
  $addin.AutomationSimilarSelect() | Out-Null
  $snap = SnapAt '宋体10.5pt'
  $fails += Check $snap '五号' 'C 10.5pt 应显示中文「五号」'
  $addin.AutomationCloseSession() | Out-Null
  Start-Sleep -Milliseconds 400

  # ── D：12.5pt 非档位 → 仍显示数值（不能硬套中文名）──
  BuildSample '宋体' 12.5 3 $false
  $word.Selection.SetRange(1, 1)
  $addin.AutomationSimilarSelect() | Out-Null
  $snap = SnapAt '宋体12.5pt'
  $fails += Check $snap '12.5' 'D 非档位 12.5 应显示数值'
  $fails += CheckNot $snap '号' 'D 非档位不得显示中文号数'
  $addin.AutomationCloseSession() | Out-Null
  Start-Sleep -Milliseconds 400

  # ── E：黑体 16pt → 「三号」──
  BuildSample '黑体' 16 3 $false
  $word.Selection.SetRange(1, 1)
  $addin.AutomationSimilarSelect() | Out-Null
  $snap = SnapAt '黑体16pt'
  $fails += Check $snap '三号' 'E 16pt 应显示中文「三号」'
  $addin.AutomationCloseSession() | Out-Null
  Start-Sleep -Milliseconds 400

  # ── F：文本同类模式 → 字号不得为 0（回归）──
  BuildSample '宋体' 14 3 $true
  $txt = $doc.Content.Text
  $idx = $txt.IndexOf('正确')
  if ($idx -ge 0) {
    $word.Selection.SetRange(1 + $idx, 1 + $idx + 2)
    $addin.AutomationTextSimilarSelect() | Out-Null
    $snap = SnapAt '文本同类'
    $fails += CheckNot $snap '|0' 'F 文本同类字号不得为 0'
    $fails += Check $snap '四号' 'F 文本同类应显示命中处真实字号'
    $addin.AutomationCloseSession() | Out-Null
  } else { Write-Host "  [SKIP] F 无「正确」样本" }

  if ($fails -eq 0) { Write-Host "  浮条显示取证：ALL PASS" } else { Write-Host "  浮条显示取证：FAIL $fails" }
}
catch {
  Write-Host "  [ERROR] $($_.Exception.Message)"
  Write-Host $_.ScriptStackTrace
  $fails = 99
}
finally {
  try { if ($doc) { $doc.Close(0) } } catch { }
  try { if ($word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } } catch { }
}
exit $fails
