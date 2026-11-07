# A/B：OMaths 索引器 vs 枚举器（同文档）
$ErrorActionPreference = 'Continue'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
Start-Sleep -Seconds 5
$doc = $word.Documents.Open('E:\TypeRight\tests\第二十六章 二次函数-学生.docx', $null, $true)
Start-Sleep -Seconds 2
$count = $doc.OMaths.Count
# A. 索引器（v0.2.2 之前的旧路径）
$t0 = Get-Date
$zonesA = 0
for ($i = 1; $i -le $count; $i++) { $r = $doc.OMaths.Item($i).Range; $s = $r.Start; $e = $r.End; $zonesA++ }
$t1 = Get-Date
Write-Host ("A 索引器 {0} 个: {1}ms" -f $zonesA, [math]::Round(($t1-$t0).TotalMilliseconds))
# B. 枚举器（当前兼容路径）
$t0 = Get-Date
$zonesB = 0
foreach ($m in $doc.OMaths) { $s = $m.Range.Start; $e = $m.Range.End; $zonesB++ }
$t1 = Get-Date
Write-Host ("B 枚举器 {0} 个: {1}ms" -f $zonesB, [math]::Round(($t1-$t0).TotalMilliseconds))
# C. 一次读全体区间文本？OMaths 的 Range 拼一次读（减少调用数的思路验证）
$t0 = Get-Date
$null = $doc.OMaths.Count
$t1 = Get-Date
Write-Host ("C 仅 Count: {0}ms" -f [math]::Round(($t1-$t0).TotalMilliseconds))
$doc.Close(0)
$word.Quit()
