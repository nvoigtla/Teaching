# Verify animation click structure: on-click effect counts per slide.
# Usage: _verify_anim.ps1 [-Deck "Module 1 - Revised_test.pptx"]
param([string]$Deck = "Module 1 - Revised.pptx")
$folder = $PSScriptRoot
$deck = Join-Path $folder $Deck

# expected on-click counts, DISPLAY numbers in the 95-slide deck of
# 2026-08-27 (Nico's hand pass: 12 slides deleted, the Kroger-Albertsons
# pair adopted, the introduction slide duplicated to open the in-class
# part). Regenerated from the deck's own <p:timing> after that port and
# cross-checked against the previous canonical deck slide by slide.
$expected = @{
    2=2; 3=2; 4=3; 5=2; 6=2; 12=6; 13=3; 14=7; 16=2; 19=2; 21=3; 22=1;
    23=4; 24=2; 25=1; 26=2; 29=3; 30=3; 31=2; 32=2; 33=4; 34=1; 36=2;
    38=1; 39=2; 40=3; 41=3; 44=1; 45=2; 46=1; 48=2; 51=3; 52=4; 53=3;
    54=4; 55=1; 57=1; 58=1; 60=2; 61=2; 62=2; 63=4; 64=4; 65=4; 67=2;
    68=3; 69=1; 70=2; 71=2; 74=1; 75=2; 76=2; 78=2; 80=1; 81=1; 82=2;
    83=3; 85=4; 86=7; 87=3; 88=2; 89=2; 94=2
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
if ($bad -eq 0) { Write-Host "ALL CLICK COUNTS MATCH (63 animated slides)" }
else { Write-Host "$bad slides off" }
