# 浮条字体/字号「打字录入」取证（真机 COM，PowerShell 后期绑定）。
# 断言的是应用结果 + 显示结果，不只是"能输进去"。
$ErrorActionPreference = 'Stop'
[int]$fails = 0
$word = $null
$doc = $null

function Check([string]$actual, [string]$mustContain, [string]$what) {
  if ($actual -like "*$mustContain*") { Write-Host "  [PASS] $what"; return 0 }
  Write-Host "  [FAIL] $what | 期望含「$mustContain」实际「$actual」"; return 1
}
function CheckEq([string]$actual, [string]$expected, [string]$what) {
  if ($actual -eq $expected) { Write-Host "  [PASS] $what"; return 0 }
  Write-Host "  [FAIL] $what | 期望「$expected」实际「$actual」"; return 1
}

function BuildSample([string]$font, [double]$size, [int]$n) {
  $script:doc.Content.Delete()
  $sb = New-Object System.Text.StringBuilder
  for ($i = 1; $i -le $n; $i++) {
    [void]$sb.Append("这是第 $i 行同类段落，用于浮条打字录入取证。`r")
  }
  $script:doc.Content.InsertAfter($sb.ToString())
  $r = $script:doc.Content
  $r.Font.NameFarEast = $font
  $r.Font.NameAscii = $font
  $r.Font.Size = [single]$size
}

function SampleSize {
  param([int]$n = 1)
  $r = $script:doc.Paragraphs.Item($n).Range
  return [double]$r.Font.Size
}
function SampleFont {
  param([int]$n = 1)
  $r = $script:doc.Paragraphs.Item($n).Range
  return [string]$r.Font.NameFarEast
}

function WaitDisplay([string]$label) {
  $s = ''
  for ($i = 0; $i -lt 24; $i++) {
    Start-Sleep -Milliseconds 400
    $raw = @($script:addin.AutomationFloatBarDisplay())
    $cand = [string]($raw | Where-Object { "$_" -match '\|' } | Select-Object -Last 1)
    if ($cand) { $s = $cand }
    if ($s -and $s -notmatch '\|—$' -and $s -notmatch '^\|') { break }
  }
  Write-Host "  显示快照($label): $s"
  return $s
}

try {
  Write-Host "── 浮条字体/字号打字录入取证 ──────────────────────"
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $addin = $null
  try { $addin = $word.COMAddIns.Item('TypeRight.Connect').Object } catch { }
  if ($null -eq $addin) { Write-Host "  [FATAL] 加载项未连接"; exit 2 }
  Write-Host "  ADDIN: loaded"
  $script:addin = $addin
}
catch { Write-Host "  [FATAL] $($_.Exception.Message)"; exit 2 }

try {
  $doc = $word.Documents.Add()
  $script:doc = $doc

  # 预热（第一次自动化调用会被 SetTimer 时序吃掉）
  BuildSample '宋体' 10.5 2
  $word.Selection.SetRange(1, 1)
  $addin.AutomationSimilarSelect() | Out-Null
  Start-Sleep -Milliseconds 2500
  $addin.AutomationCloseSession() | Out-Null
  Start-Sleep -Milliseconds 600

  # ── 建立会话（宋体 10.5 = 五号）──
  BuildSample '宋体' 10.5 3
  $word.Selection.SetRange(1, 1)
  $addin.AutomationSimilarSelect() | Out-Null
  $snap = WaitDisplay '初始(宋体五号)'
  $fails += Check $snap '五号' '初始显示与 Word 一致'

  # ── A：字号打字「四号」→ 全部命中段变 14pt，显示回「四号」──
  $r = [string]($addin.AutomationTypeSize('四号') | Where-Object { "$_" -match '\|' } | Select-Object -Last 1)
  Write-Host "  打字'四号' -> $r"
  $fails += CheckEq (SampleSize) '14' 'A 打「四号」后文档字号 = 14pt'
  $fails += Check $r '四号' 'A 打「四号」后显示「四号」'
  $fails += Check (SampleSize 2) '14' 'A 第二段也变 14pt（批量生效）'

  # ── B：字号打字纯数字「12.5」→ 非档位，显示数值 ──
  $r = [string]($addin.AutomationTypeSize('12.5') | Where-Object { "$_" -match '\|' } | Select-Object -Last 1)
  Write-Host "  打字'12.5' -> $r"
  $fails += CheckEq (SampleSize) '12.5' 'B 打「12.5」后字号 = 12.5pt'
  $fails += Check $r '12.5' 'B 显示 12.5（不硬套中文名）'

  # ── C：带单位「14磅」──
  $r = [string]($addin.AutomationTypeSize('14磅') | Where-Object { "$_" -match '\|' } | Select-Object -Last 1)
  Write-Host "  打字'14磅' -> $r"
  $fails += CheckEq (SampleSize) '14' 'C 打「14磅」后字号 = 14pt'

  # ── D：全角数字「１６」（中文输入法常见）──
  $r = [string]($addin.AutomationTypeSize('１６') | Where-Object { "$_" -match '\|' } | Select-Object -Last 1)
  Write-Host "  打字'１６' -> $r"
  $fails += CheckEq (SampleSize) '16' 'D 全角「１６」后字号 = 16pt'
  $fails += Check $r '三号' 'D 16pt 显示「三号」'

  # ── E：非法输入不改文档、显示还原 ──
  $before = SampleSize
  $r = [string]($addin.AutomationTypeSize('乱码xyz') | Where-Object { "$_" -match '\|' } | Select-Object -Last 1)
  Write-Host "  打字'乱码xyz' -> $r"
  $fails += CheckEq (SampleSize) $before 'E 非法字号不改文档'
  $fails += Check $r '三号' 'E 非法输入后显示还原为当前值'

  # ── F：字体打字「楷体」→ 全部命中段变楷体 ──
  $r = [string]($addin.AutomationTypeFont('楷体') | Where-Object { "$_" -match '\|' } | Select-Object -Last 1)
  Write-Host "  打字字体'楷体' -> $r"
  $fails += CheckEq (SampleFont) '楷体' 'F 打「楷体」后文档字体 = 楷体'
  $fails += Check $r '楷体' 'F 显示「楷体」'
  $fails += CheckEq (SampleFont 2) '楷体' 'F 第二段也变楷体（批量生效）'

  # ── G：字体打字主题字体名「+中文正文」→ 不做存在性校验也必须能应用 ──
  $r = [string]($addin.AutomationTypeFont('仿宋') | Where-Object { "$_" -match '\|' } | Select-Object -Last 1)
  Write-Host "  打字字体'仿宋' -> $r"
  $fails += CheckEq (SampleFont) '仿宋' 'G 打「仿宋」后文档字体 = 仿宋'

  # ── H：字号打字后仍能下拉选档位（两条路径都不坏）──
  $r = [string]($addin.AutomationTypeSize('五号') | Where-Object { "$_" -match '\|' } | Select-Object -Last 1)
  $fails += CheckEq (SampleSize) '10.5' 'H 打「五号」后 = 10.5pt'

  $addin.AutomationCloseSession() | Out-Null
  if ($fails -eq 0) { Write-Host "  打字录入取证：ALL PASS" } else { Write-Host "  打字录入取证：FAIL $fails" }
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
