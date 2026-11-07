# TypeRight 发布打包：构建 → 组装离线安装包 → zip
# 用法：powershell -ExecutionPolicy Bypass -File tools\package-release.ps1
# 产物：dist\TypeRight-v<版本>.zip（解压即用，免 .NET SDK、免管理员）
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$staging = Join-Path $root 'dist\staging'
$distDir = Join-Path $root 'dist'

Write-Host '== 1/4 构建（Release publish）==' -ForegroundColor Cyan
if (Test-Path $staging) { Remove-Item -Recurse -Force $staging }
dotnet publish (Join-Path $root 'src\TypeRight.Addin\TypeRight.Addin.csproj') -c Release -o $staging
if ($LASTEXITCODE -ne 0) { throw '构建失败' }

$version = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $staging 'TypeRight.Addin.dll')).Version.ToString()
$shortVer = ($version.Split('.')[0..2] -join '.')
Write-Host "版本 v$shortVer"

Write-Host '== 2/4 组装包目录 ==' -ForegroundColor Cyan
$pkg = Join-Path $distDir "TypeRight-v$shortVer"
if (Test-Path $pkg) { Remove-Item -Recurse -Force $pkg }
New-Item -ItemType Directory -Force (Join-Path $pkg 'files') | Out-Null
New-Item -ItemType Directory -Force (Join-Path $pkg 'dist-scripts') | Out-Null

# 程序文件（全部产物）
Copy-Item (Join-Path $staging '*') (Join-Path $pkg 'files') -Force

# 脚本与说明（ps1/txt 重写为 UTF-8 BOM——中文脚本无 BOM 会被 PowerShell 当 ANSI 读成乱码）
$utf8bom = New-Object System.Text.UTF8Encoding($true)
$scriptDir = Join-Path $root 'tools\dist-scripts'
foreach ($f in @('install-core.ps1', 'repair-core.ps1', 'uninstall-core.ps1', 'install-wps-elevated.ps1')) {
    $text = [IO.File]::ReadAllText((Join-Path $scriptDir $f))
    [IO.File]::WriteAllText((Join-Path $pkg ("dist-scripts\" + $f)), $text, $utf8bom)
}
foreach ($f in @('install.bat', 'repair.bat', 'uninstall.bat')) {
    Copy-Item (Join-Path $scriptDir $f) (Join-Path $pkg $f) -Force
}
$readme = [IO.File]::ReadAllText((Join-Path $scriptDir '使用说明.txt'))
[IO.File]::WriteAllText((Join-Path $pkg '使用说明.txt'), $readme, $utf8bom)

Write-Host '== 3/4 校验关键文件 ==' -ForegroundColor Cyan
foreach ($must in @(
    (Join-Path $pkg 'files\TypeRight.Addin.dll'),
    (Join-Path $pkg 'files\TypeRight.Core.dll'),
    (Join-Path $pkg 'files\NetOffice.dll'),
    (Join-Path $pkg 'files\OfficeApi.dll'),
    (Join-Path $pkg 'files\WordApi.dll'),
    (Join-Path $pkg 'files\Newtonsoft.Json.dll'),
    (Join-Path $pkg 'install.bat'),
    (Join-Path $pkg 'repair.bat'),
    (Join-Path $pkg 'uninstall.bat'),
    (Join-Path $pkg '使用说明.txt'))) {
    if (-not (Test-Path $must)) { throw "包内缺少 $(Split-Path -Leaf $must)" }
}
$count = (Get-ChildItem (Join-Path $pkg 'files') -File).Count
Write-Host "files\ 共 $count 个文件"

Write-Host '== 4/4 压缩 ==' -ForegroundColor Cyan
$zip = Join-Path $distDir "TypeRight-v$shortVer.zip"
if (Test-Path $zip) { Remove-Item $zip }
Compress-Archive -Path (Join-Path $pkg '*') -DestinationPath $zip
$size = [math]::Round((Get-Item $zip).Length / 1MB, 2)
Write-Host ''
Write-Host "打包完成 → $zip （$size MB）" -ForegroundColor Green
Write-Host "发给朋友：整个 zip 解压后双击 install.bat 即可（无需管理员/无需 .NET SDK）。"
