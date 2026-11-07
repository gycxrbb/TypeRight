# XML 级验证：跳转后另存文档，检查 w:shd 强调色是否真实落盘
$ErrorActionPreference = 'Continue'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
Start-Sleep -Seconds 5
$src = 'E:\TypeRight\tests\初三数学代数终极讲义_一元二次方程与二次函数大一统.docx'
$copy = "$env:TEMP\jump-xml.docx"
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
$o.AutomationJumpNext()
Start-Sleep -Seconds 2
$saved = "$env:TEMP\jump-xml-out.docx"
$doc.SaveAs($saved, 12)
Start-Sleep -Seconds 1
$doc.Close(0)
$word.Quit()
Write-Host "已另存: $saved"
