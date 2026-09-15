# Export slides of an arbitrary deck in this folder to _probe\.
# Usage: _export_probe.ps1 -Deck "_probe_symbols.pptx" -Slides 1 -Width 1600
param(
    [string]$Deck = "_probe_symbols.pptx",
    [int[]]$Slides = @(1),
    [int]$Width = 1600
)
$folder = $PSScriptRoot
$path = Join-Path $folder $Deck
$dir = Join-Path $folder "_probe"
New-Item -ItemType Directory -Force $dir | Out-Null
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($path, -1, 0, 0)
$h = [int]($Width * 9 / 16)
$stem = [System.IO.Path]::GetFileNameWithoutExtension($Deck)
foreach ($i in $Slides) {
    $png = Join-Path $dir ("{0}_s{1:d2}.png" -f $stem, $i)
    $pres.Slides.Item($i).Export($png, "PNG", $Width, $h)
    Write-Host "exported $png"
}
$pres.Close()
