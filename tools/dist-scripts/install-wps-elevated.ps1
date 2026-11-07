# TypeRight WPS 支持安装（install-core.ps1 以管理员身份调起本脚本）
# 原理：WPS 是 32 位进程，读取 32 位 COM 视图（HKLM\Software\Classes\Wow6432Node）；
# 用户级 32 位注册实测不被系统解析，故 WPS 支持必须机器级注册。
$ErrorActionPreference = 'Stop'
$clsid = '{6C7C9713-51A1-4A55-8C0D-3A4E5F607182}'
$deployDir = Join-Path $env:LOCALAPPDATA 'TypeRight\Addin'
$asm = "TypeRight.Addin, Version=$($args[0]), Culture=neutral, PublicKeyToken=null"
$codeBase = Join-Path $deployDir 'TypeRight.Addin.dll'

$dst = "HKLM:\Software\Classes\Wow6432Node\CLSID\$clsid"
New-Item -Path $dst -Force | Out-Null
Set-ItemProperty $dst -Name '(default)' -Value 'TypeRight.Connect'
New-Item -Path "$dst\InprocServer32" -Force | Out-Null
New-Item -Path "$dst\ProgId" -Force | Out-Null
Set-ItemProperty "$dst\InprocServer32" -Name '(default)' -Value 'mscoree.dll'
Set-ItemProperty "$dst\InprocServer32" -Name 'ThreadingModel' -Value 'Both'
Set-ItemProperty "$dst\InprocServer32" -Name 'Assembly' -Value $asm
Set-ItemProperty "$dst\InprocServer32" -Name 'Class' -Value 'TypeRight.Addin.TypeRightAddin'
Set-ItemProperty "$dst\InprocServer32" -Name 'RuntimeVersion' -Value 'v4.0.30319'
Set-ItemProperty "$dst\InprocServer32" -Name 'CodeBase' -Value $codeBase
Set-ItemProperty "$dst\ProgId" -Name '(default)' -Value 'TypeRight.Connect'

# 自检
if (Test-Path "$dst\InprocServer32") { Write-Host 'WPS32OK'; exit 0 } else { Write-Host 'WPS32FAIL'; exit 1 }
