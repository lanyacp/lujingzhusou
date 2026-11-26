# Real-mouse clicker for the add-in panel (DPI-aware, physical pixels).
# Finds the panel window (class ThunderDFrame) of the running PowerPoint and clicks
# the given window-relative offsets at the given times (ms since the window appears),
# then screenshots the panel so its status line can be read.
#   -Seq "2500:115,105;4500:89,153;7500:210,153"
param([string]$Seq = '2500:115,105;4500:89,153;7500:210,153', [string]$Out = 'ui-click-final.png', [int]$WaitFormSec = 40, [int]$TailMs = 2500, [switch]$ShotEach)
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class UIC {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  public static IntPtr FindForm(uint pid){
    IntPtr f=IntPtr.Zero;
    EnumWindows(delegate(IntPtr h, IntPtr l){
      uint p; GetWindowThreadProcessId(h, out p);
      if((pid==0 || p==pid) && IsWindowVisible(h)){ var c=new StringBuilder(256); GetClassName(h,c,256); if(c.ToString()=="ThunderDFrame"){ f=h; return false; } }
      return true; }, IntPtr.Zero);
    return f;
  }
  public static RECT Bring(IntPtr h){ SetWindowPos(h, (IntPtr)(-1), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040); RECT r; GetWindowRect(h, out r); return r; }
  public static void Click(int x, int y){ SetCursorPos(x, y); System.Threading.Thread.Sleep(120); mouse_event(0x0002,0,0,0,UIntPtr.Zero); System.Threading.Thread.Sleep(60); mouse_event(0x0004,0,0,0,UIntPtr.Zero); }
}
'@
[void][UIC]::SetProcessDPIAware()
$log = New-Object System.Collections.ArrayList
$w0 = [Diagnostics.Stopwatch]::StartNew()
$pp = Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -First 1   # informational only
$h = [IntPtr]::Zero
# the panel is located by window class alone (process enumeration is unreliable in this environment)
while ($h -eq [IntPtr]::Zero -and $w0.Elapsed.TotalSeconds -lt $WaitFormSec) {
  $h = [UIC]::FindForm([uint32]0)
  if ($h -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 150 }
}
if ($h -eq [IntPtr]::Zero) { Write-Host 'FORM-NOT-FOUND'; exit 1 }
$r = [UIC]::Bring($h)
[void]$log.Add('form rect=' + $r.L + ',' + $r.T + ',' + $r.R + ',' + $r.B)
$sw = [Diagnostics.Stopwatch]::StartNew()
$ci = 0
foreach ($item in ($Seq -split ';')) {
  $parts = $item -split ':'
  $at = [int]$parts[0]; $xy = $parts[1] -split ','
  while ($sw.Elapsed.TotalMilliseconds -lt $at) { Start-Sleep -Milliseconds 30 }
  [UIC]::Click($r.L + [int]$xy[0], $r.T + [int]$xy[1])
  [void]$log.Add('clicked t=' + $at + 'ms at +' + $xy[0] + ',+' + $xy[1])
  if ($ShotEach) {
    Start-Sleep -Milliseconds 700
    $ci++
    $r2 = [UIC]::Bring($h)
    $b2 = New-Object System.Drawing.Bitmap ($r2.R - $r2.L), ($r2.B - $r2.T)
    $g2 = [System.Drawing.Graphics]::FromImage($b2)
    $g2.CopyFromScreen($r2.L, $r2.T, 0, 0, (New-Object System.Drawing.Size ($r2.R - $r2.L), ($r2.B - $r2.T)))
    $g2.Dispose()
    $b2.Save((Join-Path $PSScriptRoot ('out\' + [IO.Path]::GetFileNameWithoutExtension($Out) + '_' + $ci + '.png')), [System.Drawing.Imaging.ImageFormat]::Png)
    $b2.Dispose()
    [void]$log.Add('  shot ' + $ci)
  }
}
$endAt = $sw.Elapsed.TotalMilliseconds + $TailMs
while ($sw.Elapsed.TotalMilliseconds -lt $endAt) { Start-Sleep -Milliseconds 50 }
$r = [UIC]::Bring($h)
$wd = $r.R - $r.L; $ht = $r.B - $r.T
$bmp = New-Object System.Drawing.Bitmap $wd, $ht
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($r.L, $r.T, 0, 0, (New-Object System.Drawing.Size $wd, $ht))
$g.Dispose()
$bmp.Save((Join-Path $PSScriptRoot ('out\' + $Out)), [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
[IO.File]::WriteAllLines((Join-Path $PSScriptRoot 'out\ui-click.txt'), $log, (New-Object Text.UTF8Encoding($true)))
foreach ($l in $log) { Write-Host $l }
