$ErrorActionPreference = 'Stop'
$base = 'E:\TypeRight\src\TypeRight.Addin'
Add-Type -AssemblyName PresentationCore
$asm = [Reflection.Assembly]::LoadFrom("$base\bin\Debug\TypeRight.Addin.dll")
$icons = $asm.GetType('TypeRight.Addin.UI.RibbonIcons')
$kindType = $asm.GetType('TypeRight.Addin.UI.OfficeThemeKind')
$save = $icons.GetMethod('SavePng')
$out = "$env:TEMP\typeright-icons"
New-Item -ItemType Directory -Force $out | Out-Null
foreach ($id in 'btnSimilar','btnFormulas','mnuProfiles','btnProfileSave','btnProfileManage') {
  foreach ($t in 'Dark','Light') {
    $kind = [Enum]::Parse($kindType, $t)
    $path = [string](Join-Path $out "$id-$t.png")
    $save.Invoke($null, @([string]$id, $kind, $path)) | Out-Null
  }
}
Get-ChildItem $out | Select-Object Name, Length
