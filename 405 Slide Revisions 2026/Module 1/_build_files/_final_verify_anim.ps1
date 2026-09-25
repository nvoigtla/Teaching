# Create the final slide version, step 2: COM click-count check.
# Every slide of "Module 1 - Final.pptx" must carry the same number of
# animation effects AND the same number of on-click beats as the slide it was
# copied from - the in-class block from "Module 1 - In Class.pptx", the taped
# blocks from their video deck. PowerPoint itself is the reader, which is the
# point: the XML timing diff in _final_verify.py is independent.
#
# Read-only. $folder is the module folder; set it to the parent when this
# script moves into _build_files/.
$ErrorActionPreference = "Stop"
$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-24
$vdir = Join-Path $folder "Recorded Video Slides"
$V1 = "Module 1 - Video 1 - Introduction.pptx"
$V2 = "Module 1 - Video 2 - Markets.pptx"
$V3 = "Module 1 - Video 3 - Demand and Supply.pptx"
$V4 = "Module 1 - Video 4 - Equilibrium.pptx"
$IC = "Module 1 - In Class.pptx"

# Final slide -> source deck / source slide
$deckOf = @{}
$slideOf = @{}
foreach ($i in 1..8)   { $deckOf[$i] = $V1; $slideOf[$i] = $i + 1 }
foreach ($i in 9..15)  { $deckOf[$i] = $V2; $slideOf[$i] = $i - 8 }
foreach ($i in 16..25) { $deckOf[$i] = $V3; $slideOf[$i] = $i - 15 }
foreach ($i in 26..32) { $deckOf[$i] = $V4; $slideOf[$i] = $i - 25 }
$deckOf[91] = $V1; $slideOf[91] = 11            # People Respond to Incentives
# the in-class block, in IN-CLASS order
$icPairs = @(@(34,1),@(35,2),@(36,3),@(37,4),@(38,5),@(39,6),@(40,7),@(41,8),
             @(42,9),@(43,10),@(44,11),@(45,12),@(46,13),@(47,14),@(48,15),
             @(49,16),@(50,17),@(51,18),@(52,19))
foreach ($p in $icPairs) { $deckOf[$p[0]] = $IC; $slideOf[$p[0]] = $p[1] }
foreach ($i in 53..85) { $deckOf[$i] = $IC; $slideOf[$i] = $i - 33 }
foreach ($p in @(@(87,53),@(88,54),@(89,55),@(90,56),@(92,57))) {
  $deckOf[$p[0]] = $IC; $slideOf[$p[0]] = $p[1]
}

$pp = New-Object -ComObject PowerPoint.Application
$fin = $pp.Presentations.Open((Join-Path $folder "Module 1 - Final.pptx"), -1, 0, 0)
$src = @{}
foreach ($d in @($V1, $V2, $V3, $V4)) {
  $src[$d] = $pp.Presentations.Open((Join-Path $vdir $d), -1, 0, 0)
}
$src[$IC] = $pp.Presentations.Open((Join-Path $folder $IC), -1, 0, 0)

$bad = 0
$byDeck = @{}
foreach ($i in ($deckOf.Keys | Sort-Object)) {
  $d = $deckOf[$i]
  [int]$j = $slideOf[$i]
  $a = $fin.Slides.Item($i).TimeLine.MainSequence
  $b = $src[$d].Slides.Item($j).TimeLine.MainSequence
  $ca = 0
  for ($k = 1; $k -le $a.Count; $k++) { if ($a.Item($k).Timing.TriggerType -eq 1) { $ca++ } }
  $cb = 0
  for ($k = 1; $k -le $b.Count; $k++) { if ($b.Item($k).Timing.TriggerType -eq 1) { $cb++ } }
  if (-not $byDeck.ContainsKey($d)) { $byDeck[$d] = 0 }
  if ($a.Count -ne $b.Count -or $ca -ne $cb) {
    $bad++
    Write-Output ("MISMATCH Final {0,3}: effects {1} vs {2}, clicks {3} vs {4}   <- {5} s{6}" -f `
      $i, $a.Count, $b.Count, $ca, $cb, $d, $j)
  } else {
    $byDeck[$d] = $byDeck[$d] + 1
  }
}
Write-Output ""
foreach ($d in ($byDeck.Keys | Sort-Object)) {
  Write-Output ("  {0,3} slides match  <- {1}" -f $byDeck[$d], $d)
}
Write-Output ("COM click-count check: {0} slides, {1} mismatches" -f $deckOf.Count, $bad)
foreach ($p in $src.Values) { $p.Close() }
$fin.Close()
$pp.Quit()
