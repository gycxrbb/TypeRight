# 量化：公式区枚举成本 + 段落数
$ErrorActionPreference = 'Continue'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
Start-Sleep -Seconds 5
$doc = $word.Documents.Open('E:\TypeRight\tests\第二十六章 二次函数-学生.docx', $null, $true)
Start-Sleep -Seconds 2
$t0 = Get-Date
$count = $doc.OMaths.Count
$t1 = Get-Date
Write-Host ("OMaths.Count = {0} : {1}ms" -f $count, [math]::Round(($t1-$t0).TotalMilliseconds))
$t0 = Get-Date
$n = 0
foreach ($m in $doc.OMaths) { $s = $m.Range.Start; $e = $m.Range.End; $n++ }
$t1 = Get-Date
Write-Host ("枚举 {0} 个公式+读区间: {1}ms" -f $n, [math]::Round(($t1-$t0).TotalMilliseconds))
$t0 = Get-Date
$pc = $doc.Paragraphs.Count
$t1 = Get-Date
Write-Host ("Paragraphs.Count = {0} : {1}ms" -f $pc, [math]::Round(($t1-$t0).TotalMilliseconds))
$doc.Close(0)
$word.Quit()
