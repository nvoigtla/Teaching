# Create the full slide version, step 1: assemble "Module 1 - Full.pptx" from
# "Module 1 - Revised.pptx" plus the taped decks in "Recorded Video Slides".
#
# 2026-09-24 decisions (Nico):
#   * no new "SLIDES NOT USED IN THE VIDEOS" section - Module 1's Revised deck
#     already ends with a self-contained in-class half behind its own
#     "Module 1 In-Class Part" divider, so rule 6 duplicates nothing here;
#   * rule 5 drops Video 1's own title card only - the taped order inside the
#     video block is kept, so Full opens with "Introduction", not the deck
#     title slide (which stays where it was taped, at position 7).
#
# Target: 91 slides = 92 Revised - the dropped Video 1 card.
$ErrorActionPreference = "Stop"
$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-24
$revised = Join-Path $folder "Module 1 - Revised.pptx"
$full    = Join-Path $folder "Module 1 - Full.pptx"
$vdir    = Join-Path $folder "Recorded Video Slides"
$v1 = Join-Path $vdir "Module 1 - Video 1 - Introduction.pptx"
$v2 = Join-Path $vdir "Module 1 - Video 2 - Markets.pptx"
$v3 = Join-Path $vdir "Module 1 - Video 3 - Demand and Supply.pptx"
$v4 = Join-Path $vdir "Module 1 - Video 4 - Equilibrium.pptx"
foreach ($f in @($revised, $v1, $v2, $v3, $v4)) {
  if (-not (Test-Path $f)) { throw "missing input: $f" }
}

# roll the rolling backups before the write (universal CLAUDE.md)
$t1 = Join-Path $folder "Module 1 - Full_t-1.pptx"
$t2 = Join-Path $folder "Module 1 - Full_t-2.pptx"
if (Test-Path $t1) { Copy-Item $t1 $t2 -Force }
if (Test-Path $full) { Copy-Item $full $t1 -Force }
Copy-Item $revised $full -Force

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($full, $false, $false, $true)
try {
  if ($pres.Slides.Count -ne 92) { throw "expected 92 slides in Revised, got $($pres.Slides.Count)" }

  # The taped Video 1 deck carries the BACKUP divider (R87) and "People Respond
  # to Incentives" (R91) as its slides 10-11, byte-identical to Revised's.
  # Take the taped copies so Video 1's own jump link between them survives the
  # insert, and drop the Revised originals.
  $pres.Slides(91).Delete()   # R91 People Respond to Incentives
  $pres.Slides(87).Delete()   # R87 BACKUP divider
  for ($i = 33; $i -ge 1; $i--) { $pres.Slides($i).Delete() }   # R1-R33 video part
  if ($pres.Slides.Count -ne 57) { throw "after deletion expected 57, got $($pres.Slides.Count)" }

  # taped blocks, in taped order; Video 1's slide 1 (its title card) is dropped
  $pres.Slides.InsertFromFile($v1,  0, 2, 11) | Out-Null
  $pres.Slides.InsertFromFile($v2, 10, 1,  7) | Out-Null
  $pres.Slides.InsertFromFile($v3, 17, 1, 10) | Out-Null
  $pres.Slides.InsertFromFile($v4, 27, 1,  7) | Out-Null
  if ($pres.Slides.Count -ne 91) { throw "after inserts expected 91, got $($pres.Slides.Count)" }

  # the two backup slides that rode in with Video 1 go to the Backup section,
  # at their Revised positions (rule 7: main deck -> Backup)
  $pres.Slides(10).MoveTo(90)   # People Respond to Incentives, between R90 and R92
  $pres.Slides(9).MoveTo(86)    # BACKUP divider, ahead of R88

  $pres.Save()
  Write-Output "slides: $($pres.Slides.Count)"
  for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $s = $pres.Slides($i)
    $t = ""
    foreach ($sh in $s.Shapes) {
      if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) {
        $t = ($sh.TextFrame.TextRange.Text -replace "\r|\n", " ")
        break
      }
    }
    Write-Output ("{0,3}  {1}" -f $i, $t.Substring(0, [Math]::Min(60, $t.Length)))
  }
} finally {
  $pres.Close()
  $pp.Quit()
}
