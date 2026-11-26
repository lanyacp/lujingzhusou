param([string]$Out = 'vbe-shot.png', [string]$Class = 'wndclass_desked_gsk', [string]$TitleRx = '', [switch]$NoMax)
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Runtime.InteropServices;
public static class Shot {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  public static IntPtr Found; public static string Title = "";
  public static IntPtr Find(uint pid, string cls, string rx){
    Found = IntPtr.Zero;
    EnumWindows(delegate(IntPtr h, IntPtr l){
      uint p; GetWindowThreadProcessId(h, out p);
      if(p==pid && Found==IntPtr.Zero){
        var c=new StringBuilder(256); GetClassName(h,c,256);
        var t=new StringBuilder(512); GetWindowText(h,t,512);
        bool clsOk = (cls=="") ? true : (c.ToString()==cls);
        bool vis = IsWindowVisible(h);
        if(clsOk && (rx=="" || System.Text.RegularExpressions.Regex.IsMatch(t.ToString(), rx))){
          if(vis || Found==IntPtr.Zero){ Found=h; Title=t.ToString(); if(vis) return false; }
        }
      }
      return true; }, IntPtr.Zero);
    return Found;
  }
  public static RECT Place(IntPtr h, int cmd){
    if(cmd>0) ShowWindow(h, cmd);
    if(cmd!=3) SetWindowPos(h, (IntPtr)(-1), 40, 40, 1200, 800, 0x40);
    else SetWindowPos(h, (IntPtr)(-1), 0, 0, 0, 0, 0x0001|0x0002|0x0040);
    BringWindowToTop(h); SetForegroundWindow(h);
    RECT r; GetWindowRect(h, out r); return r;
  }
  public static RECT Rect(IntPtr h){ RECT r; GetWindowRect(h, out r); return r; }
  public static void TopOnly(IntPtr h){ ShowWindow(h, 5); SetWindowPos(h, (IntPtr)(-1), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040); BringWindowToTop(h); }
}
'@
[void][Shot]::SetProcessDPIAware()
$log = New-Object System.Collections.ArrayList
$pp = Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $pp) { Write-Host 'NO-POWERPNT' }
else {
  $h = [Shot]::Find([uint32]$pp.Id, $Class, $TitleRx)
  if ($h -eq [IntPtr]::Zero) { Write-Host 'NOT-FOUND' }
  else {
    if ($NoMax) { [void][Shot]::TopOnly($h) } else { $r = [Shot]::Place($h, 3) }   # NoMax: keep natural size, just bring to front
    Start-Sleep -Milliseconds 1200
    $r = [Shot]::Rect($h)
    $w = $r.R - $r.L; $hh = $r.B - $r.T
    if ($w -lt 50 -or $hh -lt 50) { $r = [Shot]::Place($h, 9); Start-Sleep -Milliseconds 800; $r = [Shot]::Rect($h); $w = $r.R - $r.L; $hh = $r.B - $r.T }
    $bmp = New-Object System.Drawing.Bitmap $w, $hh
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($r.L, $r.T, 0, 0, (New-Object System.Drawing.Size $w, $hh))
    $g.Dispose()
    $bmp.Save((Join-Path $PSScriptRoot ('out\' + $Out)), [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    [void]$log.Add('hwnd=' + $h + ' title=[' + [Shot]::Title + '] rect=' + $r.L + ',' + $r.T + ',' + $r.R + ',' + $r.B + ' -> ' + $Out)
  }
}
[IO.File]::WriteAllLines((Join-Path $PSScriptRoot 'out\dbg-shot.txt'), $log, (New-Object Text.UTF8Encoding($true)))
foreach ($l in $log) { Write-Host $l }
