# Create the final slide version, step 2: assemble "Module 1 - Final.pptx"
# from "Module 1 - Full.pptx" plus the hand-edited "Module 1 - In Class.pptx".
#
# In Class is the source of truth for every slide it holds (rule 1) and is
# read, never written; Full is the reference for order, the video blocks and
# the backup section, and is read, never written.
#
# What this does, in order:
#   A. replaces the 12 slides that really differ (rule 3), in place;
#   B. inserts the one in-class-only slide, IC7 "Recall from Video 1:
#      Economists as Hedgehogs", at its in-class position (rule 5);
#   C. reorders the in-class block into the IN-CLASS order (rule 5): the deck
#      title slide to the front, the "Module 1 In-Class Part" divider down to
#      just before the applications divider, and the Netflix case after the
#      Kroger and Costco ones.
# Rule 4: slides 1-32, 86 and 90 came from the taped decks and are NEVER
# touched. Rule 6: Full 33 (the hidden shifts slide), Full 85 ("Next Steps")
# and Full 90 ("People Respond to Incentives") are absent from In Class and
# are KEPT at their Full positions.
#
# Target: 92 slides = 91 Full + 1 in-class-only slide.
# NOTE (2026-09-24): this reads "Module 1 - Full.pptx", which the
# step-1b cleanup deleted once Final superseded it. Full is not in git
# either - it is REGENERATED, not restored: restore
# "Module 1 - Revised.pptx" from git into the module folder, then run
# _assemble_full.ps1.
$ErrorActionPreference = "Stop"
$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-24
$full    = Join-Path $folder "Module 1 - Full.pptx"
$inclass = Join-Path $folder "Module 1 - In Class.pptx"
$final   = Join-Path $folder "Module 1 - Final.pptx"
foreach ($f in @($full, $inclass)) {
  if (-not (Test-Path $f)) { throw "missing input: $f" }
}

# roll the rolling backups before the write (universal CLAUDE.md)
$t1 = Join-Path $folder "Module 1 - Final_t-1.pptx"
$t2 = Join-Path $folder "Module 1 - Final_t-2.pptx"
if (Test-Path $t1) { Copy-Item $t1 $t2 -Force }
if (Test-Path $final) { Copy-Item $final $t1 -Force }
Copy-Item $full $final -Force

# rule 3: In-Class slide -> the Full slide it replaces, ascending.
# Each replace is one insert plus one delete, so later indices do not move.
# Kept as two int arrays: an [ordered] hashtable's keys come back typed in a
# way that PowerPoint's Slides.Item() rejects.
[int[]]$icSlide   = @( 6, 12, 16, 17, 18, 30, 32, 34, 40, 45, 46, 52)
[int[]]$fullSlide = @(40, 45, 48, 50, 51, 62, 64, 66, 72, 77, 78, 84)

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($final, $false, $false, $true)
try {
  if ($pres.Slides.Count -ne 91) { throw "expected 91 slides in Full, got $($pres.Slides.Count)" }

  # ---- A. replace the changed slides -------------------------------------
  for ($k = 0; $k -lt $icSlide.Length; $k++) {
    [int]$ic = $icSlide[$k]
    [int]$at = $fullSlide[$k]
    # InsertFromFile puts the copy AFTER $at, so the old slide is then at $at
    $pres.Slides.InsertFromFile($inclass, $at, $ic, $ic) | Out-Null
    $pres.Slides.Item($at).Delete()
  }
  if ($pres.Slides.Count -ne 91) { throw "after replacements expected 91, got $($pres.Slides.Count)" }

  # ---- B. the one in-class-only slide -------------------------------------
  # after "Questions and Office Hours" (40), before "Making the Most" (41)
  $pres.Slides.InsertFromFile($inclass, 40, 7, 7) | Out-Null
  if ($pres.Slides.Count -ne 92) { throw "after the new slide expected 92, got $($pres.Slides.Count)" }

  # ---- C. the in-class block into IN-CLASS order --------------------------
  $pres.Slides.Item(34).MoveTo(47)   # "Module 1 In-Class Part" divider -> before the applications divider
  $pres.Slides.Item(35).MoveTo(34)   # the deck title slide -> the front of the block
  $pres.Slides.Item(50).MoveTo(52)   # Netflix -> after Kroger-Albertsons and Costco

  $pres.Save()
  Write-Output "slides: $($pres.Slides.Count)"
  for ($i = 33; $i -le 53; $i++) {
    $s = $pres.Slides.Item($i)
    $t = ""
    foreach ($sh in $s.Shapes) {
      if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) {
        $t = ($sh.TextFrame.TextRange.Text -replace "\r|\n", " ")
        if ($t -notmatch "^Module 1 . ") { break }
      }
    }
    Write-Output ("{0,3}  {1}" -f $i, $t.Substring(0, [Math]::Min(62, $t.Length)))
  }
} finally {
  $pres.Close()
  $pp.Quit()
}
