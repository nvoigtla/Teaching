# Verify the animation click structure through PowerPoint COM: how many
# ON-CLICK effects each slide really has, against what _animate.py planned.
# Counting effects alone is not enough - a multi-paragraph range effect
# gets re-expanded by PowerPoint into separate CLICKS - so this walks
# MainSequence and counts TriggerType 1 (on click).
#
# Usage: _verify_anim.ps1 [-Deck "Module 4 - Revised.pptx"] [-Expect "..."]
#   -Expect takes the "sNN: C clicks" lines that `python _animate.py all`
#   prints, saved to a file; without it the script just reports the counts.
#
# ASCII only: PowerShell reads this file as ANSI.
param(
    [string]$Deck   = "Module 4 - Revised.pptx",
    [string]$Expect = "_anim_expected.txt"
)
$folder = $PSScriptRoot
$deckPath = Join-Path $folder $Deck
$expected = @{}
$expPath = Join-Path $folder $Expect
if (Test-Path $expPath) {
    foreach ($line in Get-Content $expPath) {
        if ($line -match '^s(\d+): (\d+) clicks') {
            $expected[[int]$matches[1]] = [int]$matches[2]
        }
    }
    Write-Host ("expectations loaded for {0} slides" -f $expected.Count)
}

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($deckPath, -1, 0, 0)
$bad = 0
$anim = 0
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $seq = $pres.Slides.Item($i).TimeLine.MainSequence
    $clicks = 0
    for ($j = 1; $j -le $seq.Count; $j++) {
        if ($seq.Item($j).Timing.TriggerType -eq 1) { $clicks++ }
    }
    if ($clicks -gt 0) { $anim++ }
    if ($expected.Count -gt 0) {
        if ($expected.ContainsKey($i)) {
            if ($clicks -ne $expected[$i]) {
                Write-Host ("s{0}: MISMATCH got {1} expected {2}" -f $i, $clicks, $expected[$i])
                $bad++
            }
        } elseif ($clicks -gt 0) {
            Write-Host ("s{0}: UNEXPECTED animation ({1} clicks)" -f $i, $clicks)
            $bad++
        }
    } else {
        Write-Host ("s{0}: {1} clicks" -f $i, $clicks)
    }
}
Write-Host ("slides: {0}  animated: {1}  mismatches: {2}" -f $pres.Slides.Count, $anim, $bad)
$pres.Close()
$pp.Quit()
