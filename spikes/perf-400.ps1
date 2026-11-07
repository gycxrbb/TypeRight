# 性能验证：400 段同类 → SimilarSelect → 连续 30 次光标移动计时（旧代码在此规模每次移动 ~800 次 COM 调用）
$ErrorActionPreference = 'Continue'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
Start-Sleep -Seconds 6
$doc = $word.Documents.Add()
Start-Sleep -Seconds 2
# 400 段同格式正文
$sb = New-Object System.Text.StringBuilder
for ($i = 1; $i -le 400; $i++) { [void]$sb.Append("第 $i 段：一元二次方程的解法与判别式分析，本段与全文各段格式完全一致用于性能压测。`r") }
$doc.Range(0, 0).Text = $sb.ToString()
Start-Sleep -Seconds 2
$paras = $doc.Paragraphs
Write-Host "文档段落数: $($paras.Count)"
$p50 = $paras.Item(50)
$word.Selection.SetRange($p50.Range.Start, $p50.Range.Start)
Start-Sleep -Seconds 1
$o = $word.COMAddIns.Item('TypeRight.Connect').Object
$o.AutomationSimilarSelect()
Start-Sleep -Seconds 8
# 光标在命中段之间快速跳动 30 次（每次触发 WindowSelectionChange → HandleSelectionChange）
$sw = [System.Diagnostics.Stopwatch]::StartNew()
for ($i = 1; $i -le 30; $i++) {
  $idx = 30 + ($i * 11) % 340
  $pp = $paras.Item($idx)
  try { $word.Selection.SetRange($pp.Range.Start, $pp.Range.Start) } catch {}
}
$sw.Stop()
Write-Host ("30 次命中段内光标移动（400 命中规模）: {0:N0} ms 总 / {1:N2} ms 每次" -f $sw.Elapsed.TotalMilliseconds, ($sw.Elapsed.TotalMilliseconds / 30))
# 会话关闭（移出命中段）
$pLast = $paras.Item($paras.Count)
$word.Selection.SetRange($pLast.Range.Start, $pLast.Range.Start)
Start-Sleep -Seconds 2
$doc.Close(0)
$word.Quit()
Write-Host "完成"
