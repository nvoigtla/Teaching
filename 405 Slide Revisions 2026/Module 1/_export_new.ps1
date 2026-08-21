# Export the rebuilt deck's slides to PNGs (read-only COM, never Quit).
$folder = "c:\Users\nvoigtla\Claude Code\Teaching\405 Slide Revisions 2026\Module 1"
$deck = Join-Path $folder "Module 1 - Revised.pptx"
$dir = Join-Path $folder "_renders_new"
if (Test-Path $dir) { Remove-Item -Recurse -Force $dir }
New-Item -ItemType Directory -Force $dir | Out-Null

$pp = New-Object -ComObject PowerPoint.Application
try {
    $pres = $pp.Presentations.Open($deck, -1, 0, 0)
} catch {
    Write-Host "OPEN FAILED: $($_.Exception.Message)"
    exit 1
}
$count = $pres.Slides.Count
$ok = 0
for ($i = 1; $i -le $count; $i++) {
    $png = Join-Path $dir ("s{0:d2}.png" -f $i)
    try {
        $pres.Slides.Item($i).Export($png, "PNG", 640, 360)
        $ok++
    } catch {
        Write-Host "EXPORT FAILED s$($i): $($_.Exception.Message)"
    }
}
$pres.Close()
Write-Host "$ok / $count exported (deck opens in PowerPoint)"
