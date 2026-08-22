# Slideshow-renderer probe.  The editing canvas, PNG export and COM slideshow
# stepping can all pass while the real full-screen Slide Show fails ("The slide
# failed to open properly") -- add-ins scan the WHOLE deck at show start, so a
# broken poll slide surfaces the banner on whatever slide is showing.  This runs
# the real show and screenshots the screenClass window via PrintWindow.
param(
  [string]$Deck = "c:\Users\nvoigtla\Claude Code\Teaching\Italy IBR\Class 2\Class 2 - Revised.pptx",
  [string]$Out  = "c:\Users\nvoigtla\Claude Code\Teaching\Italy IBR\Class 2\_probe",
  [int[]]$Slides = @(1,4,7,16)
)

Add-Type -TypeDefinition @'
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
public class Shot {
  [DllImport("user32.dll")] public static extern IntPtr FindWindow(string c, string n);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  public struct RECT { public int L, T, R, B; }
  public static string Grab(string path) {
    IntPtr h = FindWindow("screenClass", null);
    if (h == IntPtr.Zero) return "NO_WINDOW";
    RECT r; GetWindowRect(h, out r);
    int w = r.R - r.L, ht = r.B - r.T;
    if (w <= 0 || ht <= 0) return "BAD_RECT";
    using (Bitmap bm = new Bitmap(w, ht))
    using (Graphics g = Graphics.FromImage(bm)) {
      IntPtr dc = g.GetHdc();
      PrintWindow(h, dc, 2);          // flag 2 = render DirectX content
      g.ReleaseHdc(dc);
      bm.Save(path, ImageFormat.Png);
    }
    return "OK " + w + "x" + ht;
  }
}
'@ -ReferencedAssemblies System.Drawing

New-Item -ItemType Directory -Force $Out | Out-Null
$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $pp.Presentations.Open($Deck, $false, $false, $true)
$win = $pres.SlideShowSettings.Run()
Start-Sleep -Seconds 4
foreach ($i in $Slides) {
  $win.View.GotoSlide($i)
  Start-Sleep -Milliseconds 1400
  $res = [Shot]::Grab(("{0}\show_s{1:d2}.png" -f $Out, $i))
  Write-Output ("slide {0,2}: {1}" -f $i, $res)
}
$win.View.Exit()
$pres.Close(); $pp.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
Write-Output "probe done"
