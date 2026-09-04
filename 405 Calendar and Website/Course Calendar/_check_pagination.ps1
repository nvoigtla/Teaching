# Verify that every week of the course calendar fits on ONE page
# (2026-08-28, Nico). Word is the only thing on this machine that knows
# where the page breaks actually fall, so drive it over COM.
#
# Usage:
#   powershell -File _check_pagination.ps1                     # canonical deck
#   powershell -File _check_pagination.ps1 -Docx "path\to.docx"
#
# Exits 1 and lists the offenders if any week spans more than one page.

param(
    [string]$Docx = "$PSScriptRoot\Calendar EMBA Hybrid -- Fall 2026.docx"
)

if (-not (Test-Path $Docx)) { Write-Error "not found: $Docx"; exit 1 }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open((Resolve-Path $Docx).Path, $false, $true)

$total = $doc.ComputeStatistics(2)   # wdStatisticPages

# page each Week bookmark starts on, in week order
$starts = @{}
foreach ($bm in $doc.Bookmarks) {
    if ($bm.Name -match '^Week(\d+)$') {
        $starts[[int]$Matches[1]] = $bm.Range.Information(3)
    }
}

$doc.Close(0)
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null

$nums = $starts.Keys | Sort-Object
$bad = @()
foreach ($n in $nums) {
    # a week ends where the next one starts; the last ends at the last page
    $next = $starts[$n + 1]
    if ($null -eq $next) { $end = $total } else { $end = $next - 1 }
    $span = $end - $starts[$n] + 1
    $mark = if ($span -eq 1) { "ok" } else { "SPANS $span PAGES" }
    "Week {0,-2}  page {1,-3} {2}" -f $n, $starts[$n], $mark
    if ($span -ne 1) { $bad += $n }
}

""
"total pages: $total"
if ($bad.Count -gt 0) {
    Write-Error ("weeks on more than one page: " + ($bad -join ", "))
    exit 1
}
"PASS - every week fits on one page"
