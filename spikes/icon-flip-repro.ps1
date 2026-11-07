# 直接 Invalidate（等价于主题翻转后 RefreshRibbonTheme 的动作）
$o = $word.COMAddIns.Item('TypeRight.Connect').Object
$o.AutomationRefreshRibbon()
Start-Sleep -Seconds 4
ShotTop $proc.MainWindowHandle 'E:\TypeRight\spikes\flip-after-invalidate.png'
$doc.Close(0)
$word.Quit()
Write-Host "── GetImage 日志 ──"
Get-Content "$env:LOCALAPPDATA\TypeRight\addin-log.txt" -Tail 20 | Select-String -Pattern 'GetImage'
