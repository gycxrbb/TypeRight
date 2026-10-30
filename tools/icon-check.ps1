$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$dir = "$env:TEMP\typeright-icons"
$spec = @{
  'btnSimilar'       = 32
  'btnFormulas'      = 32
  'mnuProfiles'      = 32
  'btnProfileSave'   = 16
  'btnProfileManage' = 16
}
foreach ($name in $spec.Keys) {
  foreach ($theme in 'Dark','Light') {
    $bmp = New-Object System.Drawing.Bitmap("$dir\$name-$theme.png")
    $w = $bmp.Width; $h = $bmp.Height
    if ($w -ne $spec[$name] -or $h -ne $spec[$name]) { Write-Output "FAIL $name-$theme size ${w}x${h}"; continue }
    $visible = 0; $sumR = 0; $sumG = 0; $sumB = 0
    $minX = 999; $maxX = -1; $minY = 999; $maxY = -1
    for ($y = 0; $y -lt $h; $y++) {
      for ($x = 0; $x -lt $w; $x++) {
        $c = $bmp.GetPixel($x, $y)
        if ($c.A -gt 32) {
          $visible++
          $sumR += $c.R; $sumG += $c.G; $sumB += $c.B
          if ($x -lt $minX) { $minX = $x }; if ($x -gt $maxX) { $maxX = $x }
          if ($y -lt $minY) { $minY = $y }; if ($y -gt $maxY) { $maxY = $y }
        }
      }
    }
    $bmp.Dispose()
    $cover = [math]::Round(100.0 * $visible / ($w * $h), 1)
    $lum = if ($visible -gt 0) { [math]::Round(($sumR * 0.299 + $sumG * 0.587 + $sumB * 0.114) / $visible, 0) } else { -1 }
    $status = 'OK'
    if ($cover -lt 4 -or $cover -gt 55) { $status = 'FAIL cover' }
    if ($minX -lt 1 -or $minY -lt 1 -or $maxX -gt ($w - 2) -or $maxY -gt ($h - 2)) { $status = 'FAIL clipped' }
    if ($theme -eq 'Dark' -and $lum -lt 120) { $status = 'FAIL lum(dark wants light strokes)' }
    if ($theme -eq 'Light' -and $lum -gt 130) { $status = 'FAIL lum(light wants dark strokes)' }
    Write-Output ("{0,-9} {1,-5} cover={2,5}%  lum={3,3}  box=({4},{5})-({6},{7})  {8}" -f $name, $theme, $cover, $lum, $minX, $minY, $maxX, $maxY, $status)
  }
}
