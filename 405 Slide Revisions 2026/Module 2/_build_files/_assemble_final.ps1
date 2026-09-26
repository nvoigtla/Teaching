# Create the final slide version, step 2: assemble "Module 2 - Final.pptx"
# from "Module 2 - Full.pptx" plus the hand-edited "Module 2 - In Class.pptx".
#
# In Class is the source of truth for every slide it holds (rule 1) and is
# read, never written; Full is the reference for the video blocks and is
# read, never written.
#
# 2026-09-25 decisions (Nico):
#   * step 2 applies to Full 1-71 only; Full 72-116 (Videos 1-3) are FROZEN
#     (rule 4) and come across untouched;
#   * each old two-slide PollEv pair in Full (instructions + chart:
#     4+5, 11+12, 45+46, 57+58, 64+65; single 32, 36) is REPLACED by the one
#     newly embedded poll slide In Class carries for that activity;
#   * the placeholder "Poll: Mega Millions Elasticity" (Full 40) is replaced
#     by In Class's live Mega Millions poll (IC36);
#   * "Logistics" (Full 2) and "Recap of Module 1" (Full 3), absent from In
#     Class, are DROPPED;
#   * no in-class applications section: every in-class slide is in In Class.
# So every Full slide 1-71 is either replaced or dropped, and the in-class
# block of Final is In Class 1-64 in IN-CLASS order. There are no internal
# jump links in either deck, so nothing needs re-pointing.
#
# Target: 109 slides = In Class 64 + Full 72-116 (45).
# NOTE (2026-09-25): "Module 2 - Full.pptx" and "Module 2 - Revised.pptx"
# were deleted by the step-1b cleanup. Full was a byte copy of Revised, so
# to use this again restore Revised from git and copy it:
#   git checkout b5484ae5 -- "405 Slide Revisions 2026/Module 2/Module 2 - Revised.pptx"
#   copy "Module 2 - Revised.pptx" "Module 2 - Full.pptx"
$ErrorActionPreference = "Stop"
$folder  = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-25
$full    = Join-Path $folder "Module 2 - Full.pptx"
$inclass = Join-Path $folder "Module 2 - In Class.pptx"
$final   = Join-Path $folder "Module 2 - Final.pptx"
foreach ($f in @($full, $inclass)) {
  if (-not (Test-Path $f)) { throw "missing input: $f" }
}

# roll the rolling backups before the write (universal CLAUDE.md)
$t1 = Join-Path $folder "Module 2 - Final_t-1.pptx"
$t2 = Join-Path $folder "Module 2 - Final_t-2.pptx"
if (Test-Path $t1) { Copy-Item $t1 $t2 -Force }
if (Test-Path $final) { Copy-Item $final $t1 -Force }
Copy-Item $full $final -Force

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($final, $false, $false, $true)
try {
  if ($pres.Slides.Count -ne 116) { throw "expected 116 slides in Full, got $($pres.Slides.Count)" }
  for ($i = 71; $i -ge 1; $i--) { $pres.Slides.Item($i).Delete() }
  if ($pres.Slides.Count -ne 45) { throw "after deletion expected 45, got $($pres.Slides.Count)" }
  $pres.Slides.InsertFromFile($inclass, 0, 1, 64) | Out-Null
  if ($pres.Slides.Count -ne 109) { throw "after insert expected 109, got $($pres.Slides.Count)" }
  $pres.Save()
  Write-Output "slides: $($pres.Slides.Count)"
} finally {
  $pres.Close()
  $pp.Quit()
}
