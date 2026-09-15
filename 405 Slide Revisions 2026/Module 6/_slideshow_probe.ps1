# Full-screen slideshow probe: run the real Slide Show and capture the
# screenClass window via PrintWindow on selected slides.
param([int[]]$Slides = @(1, 21, 32, 39, 52, 66, 70, 83),
      [string]$Deck = "Module 4 - Revised.pptx")
$folder = $PSScriptRoot
$deck = Join-Path $folder $Deck
$outDir = Join-Path $folder "_probe"
New-Item -ItemType Directory -Force $outDir | Out-Null

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public class Win32Probe {
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lp);
    [DllImport("user32.dll")] public static extern int GetClassName(IntPtr hWnd, StringBuilder name, int count);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    public struct RECT { public int Left, Top, Right, Bottom; }
    public static IntPtr found = IntPtr.Zero;
    public static uint targetPid = 0;
    public static bool Cb(IntPtr h, IntPtr l) {
        var sb = new StringBuilder(256);
        GetClassName(h, sb, 256);
        if (sb.ToString() == "screenClass") {
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (targetPid == 0 || pid == targetPid) { found = h; return false; }
        }
        return true;
    }
    public static IntPtr FindShow(uint pid) {
        found = IntPtr.Zero; targetPid = pid;
        EnumWindows(Cb, IntPtr.Zero);
        return found;
    }
}
"@ -ReferencedAssemblies System.Drawing

Add-Type -AssemblyName System.Drawing

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = -1
$pres = $pp.Presentations.Open($deck, -1, 0, -1)
$pid2 = (Get-Process POWERPNT | Select-Object -First 1).Id
$show = $pres.SlideShowSettings.Run()
# wait until the show is actually navigable (fresh instances + the PollEv
# add-in's deck scan can take a while)
$ready = $false
for ($t = 0; $t -lt 30; $t++) {
    Start-Sleep -Seconds 2
    try { $show.View.GotoSlide(1); $ready = $true; break } catch {}
}
if (-not $ready) { Write-Host "WARN: slideshow never became navigable" }

$hwnd = [Win32Probe]::FindShow(0)
if ($hwnd -eq [IntPtr]::Zero) {
    Write-Host "NO screenClass window found"
} else {
    foreach ($n in $Slides) {
        try { $show.View.GotoSlide($n) } catch { Write-Host "goto $n failed: $($_.Exception.Message)"; continue }
        Start-Sleep -Milliseconds 2500
        $r = New-Object Win32Probe+RECT
        [Win32Probe]::GetWindowRect($hwnd, [ref]$r) | Out-Null
        $w = $r.Right - $r.Left; $h = $r.Bottom - $r.Top
        if ($w -le 0 -or $h -le 0) { Write-Host "s$($n): zero-size window"; continue }
        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $hdc = $g.GetHdc()
        [Win32Probe]::PrintWindow($hwnd, $hdc, 2) | Out-Null
        $g.ReleaseHdc($hdc); $g.Dispose()
        $small = New-Object System.Drawing.Bitmap($bmp, 640, [int](640 * $h / $w))
        $png = Join-Path $outDir ("probe_s{0:d2}.png" -f $n)
        $small.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose(); $small.Dispose()
        Write-Host "s$($n): captured"
    }
}
try { $show.View.Exit() } catch {}
Start-Sleep -Seconds 1
$pres.Close()
Write-Host "probe done"
