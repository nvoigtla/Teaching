param([string]$Deck, [string]$Slides, [string]$OutDir)
$ErrorActionPreference = "Stop"
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 400
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open((Resolve-Path $Deck).Path, $true, $false, $false)
foreach ($n in $Slides.Split(",")) {
  $i = [int]$n
  $p = Join-Path (Resolve-Path $OutDir).Path ("s{0:d2}.png" -f $i)
  $pres.Slides.Item($i).Export($p, "PNG", 1400, 788)
  Write-Output ("exported " + $p)
}
$pres.Close(); $app.Quit()
