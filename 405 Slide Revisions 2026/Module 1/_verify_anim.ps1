# Verify animation click structure: on-click effect counts per slide.
# Usage: _verify_anim.ps1 [-Deck "Module 1 - Revised_test.pptx"]
param([string]$Deck = "Module 1 - Revised.pptx")
$folder = $PSScriptRoot
$deck = Join-Path $folder $Deck

# Expected on-click counts, DISPLAY numbers in the 92-slide deck of
# 2026-09-23 (the In-Class merge). Regenerated from the deck's own
# <p:timing>: this deck is HAND-EDITED and is the source of truth, so
# this is a regression check - it tells you an animation was lost, e.g.
# by a stray PowerPoint save, not that a rebuild matches a plan.
$expected = @{
    2=2; 3=2; 4=3; 5=2; 6=2; 12=6; 13=3; 14=7; 16=2; 19=2; 21=3; 22=1;
    23=4; 24=2; 25=1; 26=2; 29=3; 30=3; 31=2; 32=2; 33=4; 34=1; 36=2;
    38=1; 39=2; 40=2; 41=3; 43=1; 44=1; 45=1; 46=2; 49=3; 50=3; 51=3;
    52=4; 53=1; 55=1; 56=1; 58=2; 59=1; 60=2; 61=4; 62=3; 63=4; 65=2;
    66=10; 67=1; 68=1; 69=2; 71=1; 72=2; 73=2; 75=2; 77=1; 78=1; 79=3;
    80=3; 82=6; 83=23; 84=3; 85=2; 86=1; 91=2; 92=13
}

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($deck, -1, 0, 0)
$bad = 0
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $seq = $pres.Slides.Item($i).TimeLine.MainSequence
    $clicks = 0
    for ($j = 1; $j -le $seq.Count; $j++) {
        if ($seq.Item($j).Timing.TriggerType -eq 1) { $clicks++ }
    }
    if ($expected.ContainsKey($i)) {
        if ($clicks -ne $expected[$i]) {
            Write-Host "s$($i): MISMATCH got $clicks expected $($expected[$i])"
            $bad++
        }
    } elseif ($clicks -gt 0) {
        Write-Host "s$($i): UNEXPECTED animation ($clicks clicks)"
        $bad++
    }
}
Write-Host "slides: $($pres.Slides.Count)"
$pres.Close()
if ($bad -eq 0) { Write-Host "ALL CLICK COUNTS MATCH (64 animated slides)" }
else { Write-Host "$bad slides off" }
