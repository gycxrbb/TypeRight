# TypeRight 安装脚本
# 管理员：RegAsm（HKLM，NetOffice 钩子自动写 Word\Addins）+ 补齐 ProgId/InprocServer32 键
# 非管理员：HKCU 手工注册（免管理员，Word 以当前用户身份加载）
# 用法：powershell -ExecutionPolicy Bypass -File tools\install.ps1
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$deployDir = Join-Path $env:LOCALAPPDATA "TypeRight\Addin"
$addinClsid = "{6C7C9713-51A1-4A55-8C0D-3A4E5F607182}"     # TypeRight.Connect
$paneClsid  = "{B3E4D2A1-9F56-4C71-8D2E-6A7B8C9D0E1F}"     # PreviewPaneHost（CTP 宿主）
$paneProgId = "TypeRight.Addin.Panes.PreviewPaneHost"

Write-Host "== 1/4 构建（dotnet publish）==" -ForegroundColor Cyan
dotnet publish (Join-Path $root "src\TypeRight.Addin\TypeRight.Addin.csproj") -c Release -o $deployDir
if ($LASTEXITCODE -ne 0) { throw "构建失败" }

Write-Host "== 2/4 校验产物 ==" -ForegroundColor Cyan
$dll = Join-Path $deployDir "TypeRight.Addin.dll"
if (-not (Test-Path $dll)) { throw "缺少 TypeRight.Addin.dll" }
foreach ($dep in @("TypeRight.Core.dll", "NetOffice.dll", "OfficeApi.dll", "WordApi.dll", "Newtonsoft.Json.dll")) {
    if (-not (Test-Path (Join-Path $deployDir $dep))) { throw "缺少依赖 $dep" }
}
$version = [Reflection.AssemblyName]::GetAssemblyName($dll).Version.ToString()
$assemblyName = "TypeRight.Addin, Version=$version, Culture=neutral, PublicKeyToken=null"
Write-Host "程序集版本 $version → $deployDir"

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

function Set-Reg([Microsoft.Win32.RegistryKey]$parent, [string]$name, [string]$value) {
    $parent.SetValue($name, $value)
}
function New-Sub([Microsoft.Win32.RegistryKey]$parent, [string]$name) {
    $parent.CreateSubKey($name)
}

function Register-DotNetClass([Microsoft.Win32.RegistryKey]$classesRoot, [string]$progId, [string]$clsid, [string]$classFullName, [string]$codeBase, [string]$assemblyName) {
    # ProgId -> CLSID
    $progKey = $classesRoot.CreateSubKey($progId)
    Set-Reg $progKey "" $progId
    Set-Reg (New-Sub $progKey "CLSID") "" $clsid
    # CLSID 激活信息（InprocServer32 本键必须带全部值——CoCreate 只看这里）
    $clsKey = $classesRoot.CreateSubKey("CLSID\$clsid")
    Set-Reg $clsKey "" $classFullName
    $inproc = $clsKey.CreateSubKey("InprocServer32")
    Set-Reg $inproc "" "mscoree.dll"
    Set-Reg $inproc "Assembly" $assemblyName
    Set-Reg $inproc "Class" $classFullName
    Set-Reg $inproc "RuntimeVersion" "v4.0.30319"
    Set-Reg $inproc "ThreadingModel" "Both"
    Set-Reg $inproc "CodeBase" $codeBase
    Set-Reg (New-Sub $clsKey "ProgId") "" $progId
    # 版本子键（regasm 兼容布局）
    $verKey = $inproc.CreateSubKey($version)
    Set-Reg $verKey "Assembly" $assemblyName
    Set-Reg $verKey "Class" $classFullName
    Set-Reg $verKey "RuntimeVersion" "v4.0.30319"
    Set-Reg $verKey "CodeBase" $codeBase
}

if ($isAdmin) {
    Write-Host "== 3/4 管理员模式：HKLM 注册 ==" -ForegroundColor Cyan
    $hkcu = [Microsoft.Win32.Registry]::CurrentUser
    # 清理可能残留的 HKCU 注册（避免遮蔽 HKLM，Word 失败时会把 HKCU LoadBehavior 置 2 造成遮蔽死锁）
    foreach ($p in @("Software\Classes\$paneProgId", "Software\Classes\TypeRight.Connect",
                     "Software\Classes\CLSID\$addinClsid", "Software\Classes\CLSID\$paneClsid")) {
        try { $hkcu.DeleteSubKeyTree($p) } catch { }
    }
    try { $hkcu.DeleteSubKeyTree("Software\Microsoft\Office\Word\Addins\TypeRight.Connect") } catch { }

    # 注意：传可写的 Classes 子键句柄（函数内用相对路径建键，直接传 HKLM 根会变成在根下建键）
    $hklmClasses = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('Software\Classes', $true)
    Register-DotNetClass $hklmClasses "TypeRight.Connect" $addinClsid "TypeRight.Addin.TypeRightAddin" $dll $assemblyName
    Register-DotNetClass $hklmClasses $paneProgId $paneClsid "TypeRight.Addin.Panes.PreviewPaneHost" $dll $assemblyName

    $addinKey = [Microsoft.Win32.Registry]::LocalMachine.CreateSubKey("Software\Microsoft\Office\Word\Addins\TypeRight.Connect")
    Set-Reg $addinKey "FriendlyName" "TypeRight 智能排版"
    Set-Reg $addinKey "Description" "一键智能排版：角色识别 + 方案重排 + 公式转换（NetOffice，免 VSTO 运行时）"
    Set-Reg $addinKey "LoadBehavior" 3
    Set-Reg $addinKey "CommandLineSafe" 0
}
else {
    Write-Host "== 3/4 非管理员：HKCU 注册 ==" -ForegroundColor Cyan
    $hkcuClasses = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Classes', $true)
    Register-DotNetClass $hkcuClasses "TypeRight.Connect" $addinClsid "TypeRight.Addin.TypeRightAddin" $dll $assemblyName
    Register-DotNetClass $hkcuClasses $paneProgId $paneClsid "TypeRight.Addin.Panes.PreviewPaneHost" $dll $assemblyName
    $addinKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey("Software\Microsoft\Office\Word\Addins\TypeRight.Connect")
    Set-Reg $addinKey "FriendlyName" "TypeRight 智能排版"
    Set-Reg $addinKey "Description" "一键智能排版：角色识别 + 方案重排 + 公式转换（NetOffice，免 VSTO 运行时）"
    Set-Reg $addinKey "LoadBehavior" 3
    Set-Reg $addinKey "CommandLineSafe" 0
}

# 无论管理员与否，始终补写 HKCU 注册：
# 实测本机有清理工具会持续清空 HKLM COM 注册的默认值（InprocServer32/CLSID 指向被抹，
# 附带 x-test 标记），导致 Word 不再枚举加载项；HKCU 注册不受其管辖且优先级更高。
$codeBase = Join-Path $deployDir "TypeRight.Addin.dll"
$clsidStr = "{6C7C9713-51A1-4A55-8C0D-3A4E5F607182}"
$asmName = "TypeRight.Addin, Version=$version, Culture=neutral, PublicKeyToken=null"
reg add "HKCU\Software\Classes\TypeRight.Connect" /ve /d "TypeRight.Connect" /f | Out-Null
reg add "HKCU\Software\Classes\TypeRight.Connect\CLSID" /ve /d $clsidStr /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsidStr" /ve /d "TypeRight.Connect" /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsidStr\InprocServer32" /ve /d "mscoree.dll" /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsidStr\InprocServer32" /v ThreadingModel /d "Both" /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsidStr\InprocServer32" /v Assembly /d $asmName /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsidStr\InprocServer32" /v Class /d "TypeRight.Addin.TypeRightAddin" /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsidStr\InprocServer32" /v RuntimeVersion /d "v4.0.30319" /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsidStr\InprocServer32" /v CodeBase /d $codeBase /f | Out-Null
reg add "HKCU\Software\Classes\CLSID\$clsidStr\ProgId" /ve /d "TypeRight.Connect" /f | Out-Null
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f | Out-Null
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v FriendlyName /d "TypeRight" /f | Out-Null
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v Description /d "TypeRight" /f | Out-Null
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v CommandLineSafe /t REG_DWORD /d 0 /f | Out-Null

Write-Host "== 4/4 完成 ==" -ForegroundColor Cyan
Write-Host "安装完成 → $deployDir"
Write-Host "启动 Word 后功能区将出现「智能排版」选项卡。"
Write-Host "卸载：powershell -ExecutionPolicy Bypass -File tools\uninstall.ps1（管理员模式可清理 HKLM）"
