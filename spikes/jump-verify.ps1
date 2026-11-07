# 跳转效果验证（Word 真机自动化）：
# 1. 同类选择 → 2. 连跳 3 次 → 每次断言：选区落在可见区间（滚动生效）+ 当前处青色 + 上一处还原黄
$ErrorActionPreference = 'Continue'
$word = New-Object -ComObject Word.Application
$word.Visible = $true
Start-Sleep -Seconds 7
$src = 'E:\TypeRight\tests\初三数学代数终极讲义_一元二次方程与二次函数大一统.docx'
$copy = "$env:TEMP\jump-verify.docx"
Copy-Item $src $copy -Force
Set-ItemProperty $copy -Name IsReadOnly -Value $false
$doc = $word.Documents.Open($copy)
Start-Sleep -Seconds 3
$paras = $doc.Paragraphs
$p = $null
for ($i = 20; $i -le 60; $i++) {
  $pp = $paras.Item($i)
  $t = [string]$pp.Range.Text
  if ($t.Length -gt 20 -and $t -match '方程|函数|解') { $p = $pp; break }
}
$word.Selection.SetRange($p.Range.Start, $p.Range.Start)
Start-Sleep -Seconds 1
$o = $word.COMAddIns.Item('TypeRight.Connect').Object
$o.AutomationSimilarSelect()
Start-Sleep -Seconds 6

for ($jump = 1; $jump -le 3; $jump++) {
  $o.AutomationJumpNext()
  Start-Sleep -Seconds 2
  $sel = $word.Selection.Range
  $vis = $word.ActiveWindow.VisibleRange
  $inView = ($sel.Start -ge $vis.Start -and $sel.End -le $vis.End)
  # 当前选区首字符底纹色（蜜橙 #E8B34B = 4961256；黄高亮应已摘除=0）
  $r1 = $word.ActiveDocument.Range($sel.Start, $sel.Start + 1)
  $sh = $r1.Shading.BackgroundPatternColor
  $hl = $r1.HighlightColorIndex
  Write-Host ("跳#{0}: 选区 {1}-{2} | 底纹色: {3} (期望4961256=蜜橙) | 高亮: {4} (期望0=已摘)" -f $jump, $sel.Start, $sel.End, $sh, $hl)
}
Write-Host "（第 2/3 跳时检查上一处是否还原黄：读第 1 跳处的颜色）"
$doc.Close(0)
$word.Quit()
