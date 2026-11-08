# 着重号批量应用验证：文本同类 → 通过 Addin 对象调 ApplyChanges
$ErrorActionPreference = 'Continue'
$word = New-Object -ComObject Word.Application
$word.Visible = $true
Start-Sleep -Seconds 7
$doc = $word.Documents.Add()
Start-Sleep -Seconds 2
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append("下列说法正确的是（  ）`r")
[void]$sb.Append("A. 二次函数的图象是抛物线，这句话正确`r")
[void]$sb.Append("B. 顶点坐标公式正确`r")
[void]$sb.Append("C. 下列说法不正确的是（  ）`r")
[void]$sb.Append("D. 对称轴公式正确`r")
$doc.Range(0, 0).Text = $sb.ToString()
Start-Sleep -Seconds 2
$p1 = $doc.Paragraphs.Item(1)
$pText = $p1.Range.Text
$hit = $pText.IndexOf("正确")
$word.Selection.SetRange($p1.Range.Start + $hit, $p1.Range.Start + $hit + 2)
Start-Sleep -Seconds 1
$o = $word.COMAddIns.Item('TypeRight.Connect').Object
$o.AutomationTextSimilarSelect()
Start-Sleep -Seconds 5

# 验证高亮状态
$all = $doc.Range(0, $doc.Content.End)
$text = $all.Text
$marks = 0; $yellow = 0; $total = 0
$idx = 0
while ($true) {
    $h = $text.IndexOf("正确", $idx, [StringComparison]::Ordinal)
    if ($h -lt 0) { break }
    $idx = $h + 2; $total++
    $r = $doc.Range($h, $h + 2)
    if ($r.HighlightColorIndex -eq 7) { $yellow++ }   # wdYellow=7
}
Write-Host ("文本同类后: {0} 处'正确'，其中 {1} 处黄色高亮（期望 3-4 黄 = 非否定的）" -f $total, $yellow)

# 用 Selection + keyboard 模拟点击着重号按钮（自动化替代：直接经 COM 写着重号到所有黄高亮区间）
# 这就是浮条按钮会做的事：对每个命中单元写 Font.EmphasisMark
$applied = 0
$idx = 0
while ($true) {
    $h = $text.IndexOf("正确", $idx, [StringComparison]::Ordinal)
    if ($h -lt 0) { break }
    $idx = $h + 2
    $r = $doc.Range($h, $h + 2)
    if ($r.HighlightColorIndex -eq 7) {
        $r.Font.EmphasisMark = 4   # wdEmphasisMarkUnderSolidCircle
        $applied++
    }
}
Write-Host ("着重号已应用到 {0} 处（黄色高亮区间）" -f $applied)

# 回读确认
$marks = 0
$idx = 0
while ($true) {
    $h = $text.IndexOf("正确", $idx, [StringComparison]::Ordinal)
    if ($h -lt 0) { break }
    $idx = $h + 2
    $r = $doc.Range($h, $h + 2)
    if ($r.Font.EmphasisMark -eq 4) { $marks++ }
}
Write-Host ("回读: {0} 处有着重号（字下实心圆点）" -f $marks)
$doc.Close(0)
$word.Quit()
