# Verify animation click structure: on-click effect counts per slide.
$folder = "c:\Users\nvoigtla\Claude Code\Teaching\405 Slide Revisions 2026\Module 1"
$deck = Join-Path $folder "Module 1 - Revised.pptx"

# expected on-click counts (from the _animate.py run, 2026-08-20, 87 slides)
$expected = @{
    2=2; 3=1; 4=2; 5=3; 6=3; 7=2; 8=3; 9=2; 10=3; 11=2; 12=1; 13=2; 14=1;
    18=3; 19=5; 20=4; 21=1; 23=1; 24=1; 26=2; 27=1; 28=6; 29=2; 31=2;
    32=4; 33=2; 34=4; 35=4; 36=4; 37=2; 38=4; 40=2; 41=3; 42=1; 43=2;
    44=2; 46=1; 47=2; 48=2; 50=2; 52=1; 53=1; 54=2; 55=3; 57=4; 58=7;
    59=3; 60=2; 61=2; 67=5; 68=5; 69=2; 72=2; 74=2; 75=1; 76=2; 77=2;
    78=1; 79=2; 82=3; 83=3; 84=3; 85=3; 86=4; 87=1
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
$pres.Close()
if ($bad -eq 0) { Write-Host "ALL CLICK COUNTS MATCH (65 animated slides)" }
else { Write-Host "$bad slides off" }
