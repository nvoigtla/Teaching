# "Create the full slide version" (405 Slide Revisions CLAUDE.md,
# "End of Module, step 1"):
# assemble "Module 6 - Full.pptx" from "Module 6 - Revised.pptx" and the
# taped decks in "Recorded Video Slides", driven by _full_plan.json
# (written from _full_pairing.json, which _full_inventory.py makes).
#
# PowerPoint does every copy itself (InsertFromFile / Duplicate / MoveTo),
# so notes, tags, animations, media and links travel intact.  Nothing is
# rebuilt.  Revised is only read.
$ErrorActionPreference = "Stop"
$support = $PSScriptRoot
$folder = Split-Path $PSScriptRoot -Parent   # _build_files/ since 2026-09-24
$rev = Join-Path $folder "Module 6 - Revised.pptx"
$full = Join-Path $folder "Module 6 - Full.pptx"
$vdir = Join-Path $folder "Recorded Video Slides"
$m3 = Join-Path $folder "..\Module 3\Module 3 - Revised.pptx"
$plan = Get-Content (Join-Path $support "_full_plan.json") -Raw | ConvertFrom-Json
$dot = [char]0x00B7

# rolling backups of the deliverable
$t1 = Join-Path $folder "Module 6 - Full_t-1.pptx"
$t2 = Join-Path $folder "Module 6 - Full_t-2.pptx"
if (Test-Path $t1) { Copy-Item $t1 $t2 -Force }
if (Test-Path $full) { Copy-Item $full $t1 -Force }
Copy-Item $rev $full -Force

$pp = New-Object -ComObject PowerPoint.Application
$p = $pp.Presentations.Open($full, $false, $false, $false)
if ($p.Slides.Count -ne $plan.total_rev) { throw "Revised has $($p.Slides.Count) slides, plan expects $($plan.total_rev)" }

# 1. every paired slide -> its taped version (descending, so indices hold)
foreach ($x in $plan.pairs) {
    $src = Join-Path $vdir $x.deck
    [void]$p.Slides.InsertFromFile($src, $x.r, $x.vn, $x.vn)
    $p.Slides.Item($x.r).Delete()
}
"replaced: " + $plan.pairs.Count

# slide ids by Revised number (positions are unchanged after step 1)
$id = @{}
for ($n = 1; $n -le $p.Slides.Count; $n++) { $id[$n] = $p.Slides.Item($n).SlideID }

# 2. the in-class divider, copied from Module 3 slide 93, after the
#    linked backup slide
$seatIdx = $p.Slides.FindBySlideID($id[[int]$plan.seating]).SlideIndex
[void]$p.Slides.InsertFromFile($m3, $seatIdx, 93, 93)
$div = $p.Slides.Item($seatIdx + 1)
$footer = "Management 405  $dot  Module 6  $dot  Complex Pricing and Advanced Pricing Strategies"
foreach ($sh in $div.Shapes) {
    if ($sh.HasTextFrame -and $sh.TextFrame.TextRange.Text -like "Management 405*") {
        $sh.TextFrame.TextRange.Text = $footer
    }
}
"divider at " + $div.SlideIndex

# 3. in-class copies at the end, retagged "Module 6 · In Class · Examples · <topic>"
$videoTag = "^Module 6 $dot Video \d+ $dot "
$icTag = "Module 6 $dot In Class $dot Examples $dot "
foreach ($n in $plan.copies) {
    $dup = $p.Slides.FindBySlideID($id[[int]$n]).Duplicate()
    $dup.MoveTo($p.Slides.Count)
}
# The retagging is NOT done here: editing the tag through COM makes the
# auto-fitting tag box shrink (0.42" -> 0.27", top 0 -> 0.075"), and
# switching autofit off to stop that rewrites the box.  _full_retag.py
# edits the text inside the tag's one run in the saved XML instead.
"copies: " + $plan.copies.Count

# 4. unlinked backup slides move (one copy only) to the very end
foreach ($n in $plan.move) { $p.Slides.FindBySlideID($id[[int]$n]).MoveTo($p.Slides.Count) }

# 5. drop the introduction video's title card
$p.Slides.FindBySlideID($id[[int]$plan.intro]).Delete()

$p.Save()
"slides: " + $p.Slides.Count
$p.Close()
$pp.Quit()
