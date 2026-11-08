# 仅场景 2：Esc 取消（时机提前到 600ms，覆盖注入窗口 ~0.5-1.0s）
$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class PWY {
  [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
}
"@
$word = New-Object -ComObject Word.Application
$word.Visible = $true
Start-Sleep -Seconds 7
$copy2 = "$env:TEMP\ux-cancel.docx"
Copy-Item 'E:\TypeRight\tests\初三数学代数终极讲义_一元二次方程与二次函数大一统.docx' $copy2 -Force
Set-ItemProperty $copy2 -Name IsReadOnly -Value $false
$doc2 = $word.Documents.Open($copy2)
Start-Sleep -Seconds 3
$t0 = Get-Date
$word.COMAddIns.Item('TypeRight.Connect').Object.AutomationConvertFormulas()
Start-Sleep -Milliseconds 600
[PWY]::keybd_event(0x1B, 0, 0, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 900
[PWY]::keybd_event(0x1B, 0, 2, [UIntPtr]::Zero)
Start-Sleep -Seconds 6
$z2 = $doc2.OMaths.Count
Write-Host ("取消场景: 转换后公式数 = {0}（部分值=取消生效；254=没赶上）  端到端 {1:N1}s" -f $z2, ((Get-Date) - $t0).TotalSeconds)
$doc2.Close(0)
$word.Quit()
Get-Content "$env:LOCALAPPDATA\TypeRight\addin-log.txt" -Tail 4
