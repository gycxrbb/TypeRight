# TypeRight HKCU 注册（reg.exe 直写，避开清理工具对 HKLM COM 注册的持续清除）
$ErrorActionPreference = 'Stop'
$codeBase = Join-Path $env:LOCALAPPDATA "TypeRight\Addin\TypeRight.Addin.dll"
$clsid = "{6C7C9713-51A1-4A55-8C0D-3A4E5F607182}"
$asm = "TypeRight.Addin, Version=0.1.0.0, Culture=neutral, PublicKeyToken=null"

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
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v FriendlyName /d "TypeRight" /f | Out-Null
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v Description /d "TypeRight" /f | Out-Null
reg add "HKCU\Software\Microsoft\Office\Word\Addins\TypeRight.Connect" /v CommandLineSafe /t REG_DWORD /d 0 /f | Out-Null
Write-Output "HKCU registration written via reg.exe"
