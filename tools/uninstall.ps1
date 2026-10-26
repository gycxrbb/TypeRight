# TypeRight 卸载脚本：清理 HKCU 与 HKLM（后者需管理员）的注册；部署目录保留
$addinClsid = "{6C7C9713-51A1-4A55-8C0D-3A4E5F607182}"
$paneClsid  = "{B3E4D2A1-9F56-4C71-8D2E-6A7B8C9D0E1F}"
$paneProgId = "TypeRight.Addin.Panes.PreviewPaneHost"

$roots = @([Microsoft.Win32.Registry]::CurrentUser)
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) { $roots += [Microsoft.Win32.Registry]::LocalMachine }

foreach ($root in $roots) {
    $scope = if ($root.Name -match "LocalMachine") { "HKLM" } else { "HKCU" }
    foreach ($p in @(
        "Software\Microsoft\Office\Word\Addins\TypeRight.Connect",
        "Software\Classes\TypeRight.Connect",
        "Software\Classes\$paneProgId",
        "Software\Classes\CLSID\$addinClsid",
        "Software\Classes\CLSID\$paneClsid"
    )) {
        try {
            $root.DeleteSubKeyTree($p)
            Write-Host "已删除 [$scope] $p"
        } catch { Write-Host "跳过 [$scope] $p（不存在）" }
    }
}
Write-Host "卸载完成。部署目录 %LOCALAPPDATA%\TypeRight\Addin 未删除（如需彻底清理请手动删除）。"
