# Carve the PollEverywhere slides out of "Module 4.pptx" into a small
# sidecar, using PowerPoint itself so every poll slide keeps its tags part,
# its notes part and its media.  Per Teaching CLAUDE.md: splice live content
# from a SIDECAR, never from the 15.7 MB source deck.
#
# Run once.  The sidecar is a BUILD INPUT - never delete it, and never
# round-trip it through python-pptx.
#
# ASCII only: PowerShell reads this file as ANSI, so a stray em dash breaks
# the parser.
param(
    [string]$Source = "Module 4.pptx",
    [string]$Out    = "_handoff_polls_M4.pptx",
    [int[]] $Keep   = @(20, 30, 37, 50, 62, 66)
)
$ErrorActionPreference = "Stop"
$folder = $PSScriptRoot
$src = Join-Path $folder $Source
$dst = Join-Path $folder $Out
$lock = Join-Path $folder ('~$' + $Source)

if (Test-Path $lock) { throw "Source deck is open in PowerPoint." }
Copy-Item $src $dst -Force

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($dst, $false, $false, $false)
$total = $pres.Slides.Count
for ($i = $total; $i -ge 1; $i--) {
    if ($Keep -notcontains $i) { $pres.Slides.Item($i).Delete() }
}
$pres.Save()
Write-Host ("sidecar: {0} slides kept of {1}" -f $pres.Slides.Count, $total)
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $n = $pres.Slides.Item($i).NotesPage.Shapes.Count
    Write-Host ("  slide {0}: {1} notes shapes" -f $i, $n)
}
$pres.Close()
$pp.Quit()
Write-Host ("size: {0:N1} MB" -f ((Get-Item $dst).Length / 1MB))
