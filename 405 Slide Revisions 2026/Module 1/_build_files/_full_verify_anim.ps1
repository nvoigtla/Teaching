# Create the full slide version, step 1: COM click-count check.
# Every taped slide in "Module 1 - Full.pptx" must carry the same number of
# animation effects AND the same number of on-click beats as the slide it was
# copied from in its video deck.  PowerPoint itself is the reader here, which
# is the point: the XML-level timing diff in _full_verify.py is independent.
#
# Read-only.  $folder is the module folder; set it to the parent when this
# script moves into _build_files/.
$ErrorActionPreference = "Stop"
$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-24
$vdir = Join-Path $folder "Recorded Video Slides"
$V1 = "Module 1 - Video 1 - Introduction.pptx"
$V2 = "Module 1 - Video 2 - Markets.pptx"
$V3 = "Module 1 - Video 3 - Demand and Supply.pptx"
$V4 = "Module 1 - Video 4 - Equilibrium.pptx"

# Full slide -> source deck / source slide
$deckOf = @{}
$slideOf = @{}
foreach ($i in 1..8)   { $deckOf[$i] = $V1; $slideOf[$i] = $i + 1 }  # card dropped
foreach ($i in 9..15)  { $deckOf[$i] = $V2; $slideOf[$i] = $i - 8 }
foreach ($i in 16..25) { $deckOf[$i] = $V3; $slideOf[$i] = $i - 15 }
foreach ($i in 26..32) { $deckOf[$i] = $V4; $slideOf[$i] = $i - 25 }
$deckOf[86] = $V1; $slideOf[86] = 10      # BACKUP divider
$deckOf[90] = $V1; $slideOf[90] = 11      # People Respond to Incentives

$pp = New-Object -ComObject PowerPoint.Application
$full = $pp.Presentations.Open((Join-Path $folder "Module 1 - Full.pptx"), -1, 0, 0)
$vp = @{}
foreach ($d in @($V1, $V2, $V3, $V4)) {
  $vp[$d] = $pp.Presentations.Open((Join-Path $vdir $d), -1, 0, 0)
}
$bad = 0
foreach ($i in ($deckOf.Keys | Sort-Object)) {
  $d = $deckOf[$i]
  $j = $slideOf[$i]
  $a = $full.Slides.Item($i).TimeLine.MainSequence
  $b = $vp[$d].Slides.Item($j).TimeLine.MainSequence
  $ca = 0
  for ($k = 1; $k -le $a.Count; $k++) { if ($a.Item($k).Timing.TriggerType -eq 1) { $ca++ } }
  $cb = 0
  for ($k = 1; $k -le $b.Count; $k++) { if ($b.Item($k).Timing.TriggerType -eq 1) { $cb++ } }
  if ($a.Count -ne $b.Count -or $ca -ne $cb) {
    $bad++
    Write-Output ("MISMATCH Full {0,3}: effects {1} vs {2}, clicks {3} vs {4}   <- {5} s{6}" -f `
      $i, $a.Count, $b.Count, $ca, $cb, $d, $j)
  } else {
    Write-Output ("  ok   Full {0,3}: {1} effects, {2} clicks   <- {3} s{4}" -f `
      $i, $a.Count, $ca, $d.Substring(11, 7), $j)
  }
}
Write-Output ""
Write-Output ("COM click-count check: {0} taped slides, {1} mismatches" -f $deckOf.Count, $bad)
foreach ($p in $vp.Values) { $p.Close() }
$full.Close()
$pp.Quit()
