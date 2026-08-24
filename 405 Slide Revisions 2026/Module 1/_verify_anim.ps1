# Verify animation click structure: on-click effect counts per slide.
# Usage: _verify_anim.ps1 [-Deck "Module 1 - Revised_test.pptx"]
param([string]$Deck = "Module 1 - Revised.pptx")
$folder = $PSScriptRoot
$deck = Join-Path $folder $Deck

# expected on-click counts (renumbered 2026-08-22 for the poll-pair +
# backup inserts; deck now 99 slides, same 65 animated slides)
$expected = @{
    2=2; 3=1; 4=2; 5=3; 6=3; 9=2; 10=3; 11=2; 12=3; 13=2; 14=1; 15=2;
    16=1; 20=3; 21=5; 22=4; 23=1; 26=1; 27=1; 30=2; 31=1; 32=6; 33=2;
    35=2; 36=10; 37=2; 38=4; 39=4; 40=4; 41=2; 42=4; 44=2; 45=3; 46=1;
    47=2; 48=2; 51=1; 52=2; 53=2; 55=2; 57=1; 58=1; 59=2; 60=3; 62=4;
    63=7; 64=3; 65=2; 66=2; 72=5; 73=5; 74=2; 77=2; 79=3; 80=1; 81=4;
    82=3; 83=1; 84=2; 87=3; 88=3; 89=3; 90=3; 91=4; 92=1
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
if ($bad -eq 0) { Write-Host "ALL CLICK COUNTS MATCH (65 animated slides)" }
else { Write-Host "$bad slides off" }
