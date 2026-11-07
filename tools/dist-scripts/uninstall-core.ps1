# TypeRight 卸载（uninstall.bat 调用）
$ErrorActionPreference = 'Continue'
$clsid = '{6C7C9713-51A1-4A55-8C0D-3A4E5F607182}'
$paneClsid = '{B3E4D2A1-9F56-4C71-8D2E-6A7B8C9D0E1F}'
$deployDir = Join-Path $env:LOCALAPPDATA 'TypeRight\Addin'

Write-Host ''
Write-Host '===== TypeRight 卸载 =====' -ForegroundColor Cyan

while ($true) {
    $word = Get-Process WINWORD -ErrorAction SilentlyContinue
    if (-not $word) { break }
    Write-Host '[!] 请先关闭 Word。' -ForegroundColor Yellow
    $a = Read-Host '    已关闭并继续？(y/n)'
    if ($a -eq 'n') { exit 1 }
}

reg delete "HKCU\Software\Classes\TypeRight.Connect" /f 2>$null
reg delete "HKCU\Software\Classes\CLSID\$clsid" /f 2>$null
reg delete "HKCU\Software\Classes\Wow6432Node\CLSID\$clsid" /f 2>$null
reg delete "HKCU\Software\Classes\CLSID\$paneClsid" /f 2>$null
reg delete "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /f 2>$null

# 管理员时顺带清理 HKLM（历史上管理员模式安装写入过）
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    reg delete "HKLM\Software\Classes\TypeRight.Connect" /f 2>$null
    reg delete "HKLM\Software\Classes\CLSID\$clsid" /f 2>$null
    reg delete "HKLM\Software\Classes\Wow6432Node\CLSID\$clsid" /f 2>$null
    reg delete "HKLM\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /f 2>$null
}

if (Test-Path $deployDir) {
    Remove-Item -Recurse -Force $deployDir -ErrorAction SilentlyContinue
}
Write-Host ''
Write-Host '卸载完成（注册表 + 程序文件已清除；日志目录 %LOCALAPPDATA%\TypeRight 保留，可手动删除）。' -ForegroundColor Green
