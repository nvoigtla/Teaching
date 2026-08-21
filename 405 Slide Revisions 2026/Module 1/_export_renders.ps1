# Export all slides of the Module 1 source decks to low-res PNGs.
# Attaches to the running PowerPoint (single-instance COM); opens each deck
# READ-ONLY without a window; closes only its own presentations; NEVER quits
# the application (the user may have decks open).
$folder = "c:\Users\nvoigtla\Claude Code\Teaching\405 Slide Revisions 2026\Module 1"
$outRoot = Join-Path $folder "_renders_src"
New-Item -ItemType Directory -Force $outRoot | Out-Null

$decks = @(
    @{file="Module 1 - In Class.pptx"; tag="ic"; w=600},
    @{file="Module 1 - Video 1.pptx"; tag="v1"; w=600},
    @{file="Module 1 - Video 2.pptx"; tag="v2"; w=600},
    @{file="Module 1 - Video 3.pptx"; tag="v3"; w=600},
    @{file="Module 1 - Video 4.pptx"; tag="v4"; w=600},
    @{file="Module 1 - MW.pptx"; tag="mw"; w=640}
)

$pp = New-Object -ComObject PowerPoint.Application
foreach ($d in $decks) {
    $path = Join-Path $folder $d.file
    $dir = Join-Path $outRoot $d.tag
    New-Item -ItemType Directory -Force $dir | Out-Null
    try {
        # Open(FileName, ReadOnly=-1, Untitled=0, WithWindow=0)
        $pres = $pp.Presentations.Open($path, -1, 0, 0)
    } catch {
        Write-Host "OPEN FAILED $($d.tag): $($_.Exception.Message)"
        continue
    }
    $count = $pres.Slides.Count
    $h = [int]($d.w * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
    $ok = 0
    for ($i = 1; $i -le $count; $i++) {
        $png = Join-Path $dir ("s{0:d2}.png" -f $i)
        try {
            $pres.Slides.Item($i).Export($png, "PNG", $d.w, $h)
            $ok++
        } catch {
            Write-Host "EXPORT FAILED $($d.tag) s$($i): $($_.Exception.Message)"
        }
    }
    $pres.Close()
    Write-Host "$($d.tag): $ok / $count exported"
}
Write-Host "done (PowerPoint left running)"
