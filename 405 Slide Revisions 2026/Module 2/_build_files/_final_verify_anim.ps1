# Create the final slide version, step 2: COM click-count check.
# Final 1-64 must carry the same number of animation effects AND on-click
# beats as "Module 2 - In Class.pptx" 1-64; Final 65-109 the same as their
# taped video slides. PowerPoint itself is the reader, independent of the XML
# timing diff in _final_verify.py. Read-only.
$ErrorActionPreference = "Stop"
$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-25
$vdir = Join-Path $folder "Recorded Video Slides"
$V1 = "Module 2 - Video 1 - Elasticity and Revenue.pptx"
$V2 = "Module 2 - Video 2 - Marginal Revenue.pptx"
$V3 = "Module 2 - Video 3 - Demand Estimation.pptx"
$IC = "Module 2 - In Class.pptx"

$deckOf = @{}
$slideOf = @{}
foreach ($i in 1..64)   { $deckOf[$i] = $IC; $slideOf[$i] = $i }
foreach ($i in 65..76)  { $deckOf[$i] = $V1; $slideOf[$i] = $i - 64 }
foreach ($i in 77..86)  { $deckOf[$i] = $V2; $slideOf[$i] = $i - 76 }
foreach ($i in 87..109) { $deckOf[$i] = $V3; $slideOf[$i] = $i - 86 }

$pp = New-Object -ComObject PowerPoint.Application
$fin = $pp.Presentations.Open((Join-Path $folder "Module 2 - Final.pptx"), -1, 0, 0)
Write-Output ("Final opens in PowerPoint: {0} slides" -f $fin.Slides.Count)
$src = @{}
foreach ($d in @($V1, $V2, $V3)) { $src[$d] = $pp.Presentations.Open((Join-Path $vdir $d), -1, 0, 0) }
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
