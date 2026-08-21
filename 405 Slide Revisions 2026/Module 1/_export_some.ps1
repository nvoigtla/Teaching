# Export selected slides of the rebuilt deck (pass slide numbers as args).
param([int[]]$Slides = @(12, 30, 56, 72, 75, 81, 83))
$folder = "c:\Users\nvoigtla\Claude Code\Teaching\405 Slide Revisions 2026\Module 1"
$deck = Join-Path $folder "Module 1 - Revised.pptx"
$dir = Join-Path $folder "_renders_new"
New-Item -ItemType Directory -Force $dir | Out-Null
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($deck, -1, 0, 0)
foreach ($i in $Slides) {
    $png = Join-Path $dir ("s{0:d2}.png" -f $i)
    $pres.Slides.Item($i).Export($png, "PNG", 640, 360)
}
$pres.Close()
Write-Host "exported: $($Slides -join ', ')"
