# WPS UI 验证：功能区 TypeRight 入口 + 浮动条是否出现
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

$n = (Get-Content "$env:LOCALAPPDATA\TypeRight\addin-log.txt" | Measure-Object -Line).Lines
$app = New-Object -ComObject 'KWPS.Application'
$app.Visible = $true
Start-Sleep -Seconds 8
$doc = $app.Documents.Open("$env:TEMP\wps-smoke.docx")
$doc.Activate()
Start-Sleep -Seconds 3
$p = $null
$paras = $doc.Paragraphs
for ($i = 20; $i -le 60; $i++) {
  $pp = $paras.Item($i)
  $t = [string]$pp.Range.Text
  if ($t.Length -gt 20 -and $t -match '方程|函数|解') { $p = $pp; break }
}
$app.Selection.SetRange($p.Range.Start, $p.Range.Start)
Start-Sleep -Seconds 1
$o = $app.COMAddIns.Item('TypeRight.Connect').Object
$o.AutomationSimilarSelect()
Start-Sleep -Seconds 8

Write-Host "== 1. WPS 主窗口全树搜 TypeRight 相关元素 =="
$pids = (Get-Process wps -ErrorAction SilentlyContinue).Id
$root = [System.Windows.Automation.AutomationElement]::RootElement
$children = $root.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)
foreach ($c in $children) {
  if ($pids -contains $c.Current.ProcessId) {
    Write-Host "窗口: '$($c.Current.Name)' [$($c.Current.ClassName)]"
    $all = $c.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
    Write-Host "  全树元素: $($all.Count)"
    $hit = @()
    for ($i = 0; $i -lt $all.Count; $i++) {
      $nm = $all[$i].Current.Name
      if ($nm -match 'TypeRight|同类选择|公式转换|使用说明|方案') { $hit += "$nm <$($all[$i].Current.ControlType.ProgrammaticName.Replace('ControlType.',''))>" }
    }
    Write-Host "  命中: $($hit.Count)"
    $hit | Select-Object -First 10 | ForEach-Object { Write-Host "    · $_" }
  }
}
Write-Host ""
Write-Host "== 2. 日志（浮条/会话状态）=="
Get-Content "$env:LOCALAPPDATA\TypeRight\addin-log.txt" | Select-Object -Skip $n
$doc.Close(0)
try { $app.Quit() } catch {}
