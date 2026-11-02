# TypeRight 用户侧安装核心（install.bat 调用；免管理员，HKCU 注册）
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot   # 包根目录（files\ 的上级）
$srcDir = Join-Path $root 'files'
$deployDir = Join-Path $env:LOCALAPPDATA 'TypeRight\Addin'
$clsid = '{6C7C9713-51A1-4A55-8C0D-3A4E5F607182}'

Write-Host ''
Write-Host '===== TypeRight 安装 =====' -ForegroundColor Cyan

# ── 1. .NET Framework 4.8 检查（Win10 1903+/Win11 自带；Release >= 528040）──
$release = 0
try { $release = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction Stop).Release } catch {}
if ($release -lt 528040) {
    Write-Host '[X] 缺少 .NET Framework 4.8 运行时。' -ForegroundColor Red
    Write-Host '    请安装后重试：https://dotnet.microsoft.com/download/dotnet-framework/net48' -ForegroundColor Yellow
    exit 1
}
Write-Host '[1/4] .NET Framework 4.8 ... OK' -ForegroundColor Green

# ── 2. Word 必须关闭（DLL 被占用会装一半失败）──
while ($true) {
    $word = Get-Process WINWORD -ErrorAction SilentlyContinue
    if (-not $word) { break }
    Write-Host ''
    Write-Host '[!] 检测到 Word 正在运行（' $word.Count ' 个进程），安装前需要先关闭。' -ForegroundColor Yellow
    $a = Read-Host '    已关闭 Word 并继续？(y=继续 / n=中止)'
    if ($a -eq 'n') { Write-Host '安装已中止。'; exit 1 }
}
Write-Host '[2/4] Word 已关闭 ... OK' -ForegroundColor Green

# ── 3. 复制文件 ──
if (-not (Test-Path (Join-Path $srcDir 'TypeRight.Addin.dll'))) {
    Write-Host '[X] 安装包不完整（缺 files\TypeRight.Addin.dll）。' -ForegroundColor Red; exit 1
}
New-Item -ItemType Directory -Force $deployDir | Out-Null
Copy-Item (Join-Path $srcDir '*') $deployDir -Force
$version = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $deployDir 'TypeRight.Addin.dll')).Version.ToString()
Write-Host "[3/4] 文件已复制到 $deployDir （v$version）" -ForegroundColor Green

# ── 4. HKCU 注册（免管理员；不受系统清理工具对 HKLM COM 注册的持续清除影响）──
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
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f | Out-Null
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v FriendlyName /d "TypeRight 智能排版" /f | Out-Null
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v Description /d "指认式批量排版：同类选择 + 公式转换 + 方案" /f | Out-Null
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v CommandLineSafe /t REG_DWORD /d 0 /f | Out-Null
Write-Host '[4/4] 注册完成 ... OK' -ForegroundColor Green

Write-Host ''
Write-Host '安装完成！启动（或重启）Word，功能区将出现「TypeRight」选项卡。' -ForegroundColor Green
Write-Host '若被安全软件提示，请选择"信任/允许"（未签名插件常被询问）。'
Write-Host '首次使用：光标点进段落 → 同类选择 → 浮动条改格式。详见 使用说明.txt'
