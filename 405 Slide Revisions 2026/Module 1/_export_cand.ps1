# Export the candidates deck's slides to PNGs (read-only COM, no Quit).
$folder = "c:\Users\nvoigtla\Claude Code\Teaching\405 Slide Revisions 2026\Module 1"
$deck = Join-Path $folder "Module 1 - Example Candidates.pptx"
$dir = Join-Path $folder "_renders_cand"
if (Test-Path $dir) { Remove-Item -Recurse -Force $dir }
New-Item -ItemType Directory -Force $dir | Out-Null
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($deck, -1, 0, 0)
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $pres.Slides.Item($i).Export((Join-Path $dir ("s{0:d2}.png" -f $i)), "PNG", 640, 360)
}
$c = $pres.Slides.Count
$pres.Close()
Write-Host "$c slides exported (deck opens in PowerPoint)"
