# WPS 大验证：选项卡渲染 + COMAddIns + AutomationSimilarSelect 匹配引擎
# 全部通过 powershell.exe 5.1 跑（GetActiveObject 依赖）
$ErrorActionPreference = 'Continue'

Write-Host "== 1. WPS 窗口 UIA 枚举选项卡 =="
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$pids = (Get-Process wps -ErrorAction SilentlyContinue).Id
$root = [System.Windows.Automation.AutomationElement]::RootElement
$children = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
foreach ($c in $children) {
  if ($pids -contains $c.Current.ProcessId) {
    Write-Host "WPS 窗口: '$($c.Current.Name)' [$($c.Current.ClassName)]"
    $all = $c.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
    $hit = @()
    for ($i = 0; $i -lt $all.Count; $i++) {
      $nm = $all[$i].Current.Name
      if ($nm -match 'TypeRight|同类选择|公式转换') { $hit += ("$nm <$($all[$i].Current.ControlType.ProgrammaticName)>" ) }
    }
    Write-Host "  全树元素 $($all.Count) 个，命中 TypeRight 相关: $($hit.Count)"
    $hit | Select-Object -First 8 | ForEach-Object { Write-Host "    · $_" }
  }
}

Write-Host ""
Write-Host "== 2. WPS 内 COMAddIns 名单 =="
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('KWPS.Application') }
catch { $app = New-Object -ComObject 'KWPS.Application'; $app.Visible = $true; Start-Sleep -Seconds 7 }
$n = $app.COMAddIns.Count
Write-Host "COMAddIns 数量: $n"
for ($i = 1; $i -le $n; $i++) {
  $a = $app.COMAddIns.Item($i)
  Write-Host "  $($a.ProgId)  Connect=$($a.Connect)"
}

Write-Host ""
Write-Host "== 3. 匹配引擎冒烟（AutomationSimilarSelect）=="
# 可写副本（高亮需要写权限）
$src = 'E:\TypeRight\tests\初三数学代数终极讲义_一元二次方程与二次函数大一统.docx'
$copy = "$env:TEMP\wps-smoke.docx"
Copy-Item $src $copy -Force
Set-ItemProperty $copy -Name IsReadOnly -Value $false
$doc = $app.Documents.Open($copy)
Start-Sleep -Seconds 3
$doc.Activate()
Start-Sleep -Seconds 1
Write-Host "激活文档: $($doc.Name)  活动文档名: $($app.ActiveDocument.Name)"
# 光标放一个正文段：Selection.SetRange 到「解一元二次方程的灵魂是」附近——直接用段落起点的做法
$paras = $doc.Paragraphs
Write-Host "段落数: $($paras.Count)"
# 找一个含正文的段落（跳过空段/标题），把光标放进去
$p = $null
for ($i = 20; $i -le 60; $i++) {
  $pp = $paras.Item($i)
  $t = [string]$pp.Range.Text
  if ($t.Length -gt 20 -and $t -match '方程|函数|解') { $p = $pp; Write-Host "样本段 #$i : $($t.Substring(0, [Math]::Min(24, $t.Length)))"; break }
}
if (-not $p) { $p = $paras.Item(30) }
$app.Selection.SetRange($p.Range.Start, $p.Range.Start)
Start-Sleep -Seconds 1
$addin = $app.COMAddIns.Item('TypeRight.Connect')
Write-Host "TypeRight Connect: $($addin.Connect)"
$o = $addin.Object
Write-Host "Object: $($null -ne $o)"
$mark = (Get-Content "$env:LOCALAPPDATA\TypeRight\addin-log.txt" | Measure-Object -Line).Lines
$o.AutomationSimilarSelect()
Start-Sleep -Seconds 12
Write-Host "── addin-log 新增 ──"
Get-Content "$env:LOCALAPPDATA\TypeRight\addin-log.txt" | Select-Object -Skip $mark
Write-Host ""
Write-Host "高亮命中数（黄色单元格思路→用日志 matched 判断）见上方日志"
$doc.Close(0)
