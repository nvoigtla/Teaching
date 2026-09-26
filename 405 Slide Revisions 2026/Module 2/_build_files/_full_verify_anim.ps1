# Create the full slide version, step 1: COM click-count check.
# Every video slide in "Module 2 - Full.pptx" (72-116) must carry the same
# number of animation effects AND the same number of on-click beats as its
# taped counterpart.  PowerPoint itself is the reader here, which is the
# point: the XML-level timing diff in _full_verify.py is independent.
# (Module 2: Full keeps Revised's video slides - option 1, 2026-09-25 - and
# those use Fade where the taped decks use Appear on some beats; the effect
# TYPE may differ, the counts may not.)
#
# Read-only.  $folder is the module folder.
# NOTE (2026-09-25): "Module 2 - Full.pptx" and "Module 2 - Revised.pptx"
# were deleted by the step-1b cleanup. Full was a byte copy of Revised, so
# to use this again restore Revised from git and copy it:
#   git checkout b5484ae5 -- "405 Slide Revisions 2026/Module 2/Module 2 - Revised.pptx"
#   copy "Module 2 - Revised.pptx" "Module 2 - Full.pptx"
$ErrorActionPreference = "Stop"
$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-25
$vdir = Join-Path $folder "Recorded Video Slides"
$V1 = "Module 2 - Video 1 - Elasticity and Revenue.pptx"
$V2 = "Module 2 - Video 2 - Marginal Revenue.pptx"
$V3 = "Module 2 - Video 3 - Demand Estimation.pptx"

# Full slide -> source deck / source slide
$deckOf = @{}
$slideOf = @{}
foreach ($i in 72..83)  { $deckOf[$i] = $V1; $slideOf[$i] = $i - 71 }
foreach ($i in 84..93)  { $deckOf[$i] = $V2; $slideOf[$i] = $i - 83 }
foreach ($i in 94..116) { $deckOf[$i] = $V3; $slideOf[$i] = $i - 93 }

$pp = New-Object -ComObject PowerPoint.Application
$full = $pp.Presentations.Open((Join-Path $folder "Module 2 - Full.pptx"), -1, 0, 0)
Write-Output ("Full opens in PowerPoint: {0} slides" -f $full.Slides.Count)
$vp = @{}
foreach ($d in @($V1, $V2, $V3)) {
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
Write-Output ("COM click-count check: {0} video slides, {1} mismatches" -f $deckOf.Count, $bad)
foreach ($p in $vp.Values) { $p.Close() }
$full.Close()
$pp.Quit()
