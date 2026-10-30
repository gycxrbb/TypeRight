$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$dir = "$env:TEMP\typeright-icons"
$ids = 'btnSimilar','btnFormulas','mnuProfiles','btnProfileSave','btnProfileManage'
$themes = 'Dark','Light'
$cell = 160; $pad = 24
$sheet = New-Object System.Drawing.Bitmap (($cell + $pad) * $ids.Count + $pad), (($cell + $pad) * $themes.Count + $pad)
$g = [System.Drawing.Graphics]::FromImage($sheet)
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::Half
$g.Clear([System.Drawing.Color]::FromArgb(40,40,40))
for ($ti = 0; $ti -lt $themes.Count; $ti++) {
  for ($ii = 0; $ii -lt $ids.Count; $ii++) {
    $p = Join-Path $dir "$($ids[$ii])-$($themes[$ti]).png"
    $img = [System.Drawing.Image]::FromFile($p)
    $x = $pad + $ii * ($cell + $pad); $y = $pad + $ti * ($cell + $pad)
    $g.DrawImage($img, $x, $y, $cell, $cell)
    $img.Dispose()
  }
}
$g.Dispose()
$sheet.Save("$dir\sheet.png", [System.Drawing.Imaging.ImageFormat]::Png)
$sheet.Dispose()
Write-Output "$dir\sheet.png"
