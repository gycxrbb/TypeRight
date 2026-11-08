# 字义匹配端到端验证：构造出卷场景（正确×N + 不正确×M）→ 文本同类 → 排否定 → 着重号
$ErrorActionPreference = 'Continue'
$word = New-Object -ComObject Word.Application
$word.Visible = $true
Start-Sleep -Seconds 7
$doc = $word.Documents.Add()
Start-Sleep -Seconds 2
# 构造出卷文本：6 处"正确"（含 2 处前面有否定前缀）+ 2 处"不正确"（应排除）
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append("下列说法正确的是（  ）`r")
[void]$sb.Append("A. 二次函数的图象是抛物线，这句话正确`r")
[void]$sb.Append("B. 顶点坐标是 (-b/2a, ...)，这个说法正确`r")
[void]$sb.Append("C. 下列说法不正确的是（  ）`r")   # ← 不正确，应被排除
[void]$sb.Append("D. 对称轴是 x = -b/2a，回答正确`r")
[void]$sb.Append("E. 判别式大于零有两个交点，此判断正确`r")
[void]$sb.Append("F. 未正确理解题意的同学请重读`r")   # ← 未正确，应被排除
[void]$sb.Append("G. 非正确的选项已被排除`r")       # ← 非正确，应被排除
[void]$sb.Append("答案：这道题正确答案是 A`r")
$doc.Range(0, 0).Text = $sb.ToString()
Start-Sleep -Seconds 2
$proc = Get-Process WINWORD | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1

# 光标放到第 1 行的"正确"两字内
$p1 = $doc.Paragraphs.Item(1)
$pText = $p1.Range.Text
$hit = $pText.IndexOf("正确")
$word.Selection.SetRange($p1.Range.Start + $hit, $p1.Range.Start + $hit + 2)
Start-Sleep -Seconds 1

# 文本同类（选中了"正确"两字）
$o = $word.COMAddIns.Item('TypeRight.Connect').Object
$o.AutomationTextSimilarSelect()
Start-Sleep -Seconds 6

# 验证：检查所有"正确"出现处 + 不正确处的高亮颜色
$all = $doc.Range(0, $doc.Content.End)
$text = $all.Text
$correctHits = @()
$wrongHits = @()
$idx = 0
while ($true) {
    $hit = $text.IndexOf("正确", $idx, [StringComparison]::Ordinal)
    if ($hit -lt 0) { break }
    $idx = $hit + 2
    # 读该处高亮色
    $r = $doc.Range($hit, $hit + 2)
    $hl = $r.HighlightColorIndex
    $em = $r.Font.EmphasisMark
    # 前一个字符
    $prev = if ($hit -gt 0) { $text[$hit - 1] } else { ' ' }
    $isNeg = @('不','非','未') -contains $prev
    Write-Host ("  位置{0} 前字='{1}' 排除={2} 高亮={3} 着重号={4}" -f $hit, $prev, $isNeg, $hl, $em)
    if ($isNeg) { $wrongHits += $hit } else { $correctHits += $hit }
}
Write-Host ("正确处（无否定前缀）: {0} 个；否定前缀处: {1} 个" -f $correctHits.Count, $wrongHits.Count)
$doc.Close(0)
$word.Quit()
Write-Host "── 日志 ──"
Get-Content "$env:LOCALAPPDATA\TypeRight\addin-log.txt" -Tail 4
