# TypeRight 修复（功能区选项卡消失时双击 repair.bat）
# 场景：系统清理工具抹掉注册表 / Word 崩溃后把加载项禁用（LoadBehavior 被置 2）
$ErrorActionPreference = 'Stop'
# Win7 ships PowerShell 2.0 which lacks $PSScriptRoot (PS 3.0+); old-machine fallback
if (-not $PSScriptRoot) { $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
$root = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $root 'files'
$deployDir = Join-Path $env:LOCALAPPDATA 'TypeRight\Addin'
$clsid = '{6C7C9713-51A1-4A55-8C0D-3A4E5F607182}'

Write-Host ''
Write-Host '===== TypeRight 修复 =====' -ForegroundColor Cyan

while ($true) {
    $word = Get-Process WINWORD -ErrorAction SilentlyContinue
    if (-not $word) { break }
    Write-Host '[!] 请先关闭 Word（' $word.Count ' 个进程在运行）。' -ForegroundColor Yellow
    $a = Read-Host '    已关闭并继续？(y/n)'
    if ($a -eq 'n') { exit 1 }
}

if (Test-Path (Join-Path $srcDir 'TypeRight.Addin.dll')) {
    New-Item -ItemType Directory -Force $deployDir | Out-Null
    Copy-Item (Join-Path $srcDir '*') $deployDir -Force
}
if (-not (Test-Path (Join-Path $deployDir 'TypeRight.Addin.dll'))) {
    Write-Host '[X] 找不到已安装的程序文件，请先运行 install.bat 重新安装。' -ForegroundColor Red; exit 1
}

$version = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $deployDir 'TypeRight.Addin.dll')).Version.ToString()
$asm = "TypeRight.Addin, Version=$version, Culture=neutral, PublicKeyToken=null"
$codeBase = Join-Path $deployDir 'TypeRight.Addin.dll'
reg add "HKCU\Software\Classes\TypeRight.Connect" /ve /d "TypeRight.Connect" /f | Out-Null
reg add "HKCU\Software\Classes\TypeRight.Connect\CLSID" /ve /d $clsid /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsid" /ve /d "TypeRight.Connect" /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsid\InprocServer32" /ve /d "mscoree.dll" /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsid\InprocServer32" /v ThreadingModel /d "Both" /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsid\InprocServer32" /v Assembly /d $asm /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsid\InprocServer32" /v Class /d "TypeRight.Addin.TypeRightAddin" /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsid\InprocServer32" /v RuntimeVersion /d "v4.0.30319" /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsid\InprocServer32" /v CodeBase /d $codeBase /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsid\ProgId" /ve /d "TypeRight.Connect" /f | Out-Null
# 关键一步：LoadBehavior 强制回 3（Word 禁用加载项后会写成 2，造成"选项卡消失"）
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f | Out-Null
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v FriendlyName /d "TypeRight 智能排版" /f | Out-Null

Write-Host ''
# WPS 支持补注册（仅当此前启用过：HKLM 32 位键存在时刷新版本号指向）
if (Test-Path 'HKLM:\Software\Classes\Wow6432Node\CLSID\{6C7C9713-51A1-4A55-8C0D-3A4E5F607182}') {
    try {
        $v2 = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $deployDir 'TypeRight.Addin.dll')).Version.ToString()
        Start-Process powershell -Verb RunAs -Wait -WindowStyle Hidden -ArgumentList `
            '-NoProfile','-ExecutionPolicy','Bypass','-File',`
            (Join-Path $PSScriptRoot 'install-wps-elevated.ps1'), $v2
        Write-Host 'WPS 32 位注册已刷新。' -ForegroundColor Green
    } catch { }
}

Write-Host '修复完成（注册已重建、LoadBehavior 已复位）。重启 Word 即可。' -ForegroundColor Green
