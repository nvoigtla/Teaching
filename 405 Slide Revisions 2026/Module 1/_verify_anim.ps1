# Verify animation click structure: on-click effect counts per slide.
# Usage: _verify_anim.ps1 [-Deck "Module 1 - Revised_test.pptx"]
param([string]$Deck = "Module 1 - Revised.pptx")
$folder = $PSScriptRoot
$deck = Join-Path $folder $Deck

# expected on-click counts (renumbered 2026-08-22 for the poll-pair +
# backup inserts, then the 2 Tapestry slides on 2026-08-23;
# deck is 101 slides with 67 animated.
# 2026-08-24 (Videos-Final port): 72 5->6, 84 3->2, 91/92 3->2,
# 75 now STATIC, and backup 100 gained a 2-click build.)
$expected = @{
    2=2; 3=1; 4=2; 5=3; 6=3; 9=2; 10=3; 11=2; 12=2; 13=2; 14=1; 15=2;
    16=1; 20=3; 21=5; 22=4; 23=1; 26=1; 27=1; 30=2; 31=1; 32=6; 33=2;
    35=2; 36=10; 37=2; 38=4; 39=4; 40=4; 41=2; 42=4; 44=2; 45=3; 46=1;
    47=2; 48=2; 51=1; 52=2; 53=2; 55=2; 57=1; 58=1; 59=2; 60=3; 62=4;
    63=7; 64=3; 65=2; 66=2; 72=6; 73=3; 74=7; 76=2; 79=2; 81=3;
    82=1; 83=4; 84=2; 85=1; 86=2; 89=3; 90=3; 91=2; 92=2; 93=4; 94=1;
    100=2
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
if ($bad -eq 0) { Write-Host "ALL CLICK COUNTS MATCH (67 animated slides)" }
else { Write-Host "$bad slides off" }
