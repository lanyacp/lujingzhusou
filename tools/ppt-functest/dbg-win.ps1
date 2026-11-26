param([int]$TargetPid = 0)
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class WL {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  public static List<string> Top(uint pid){
    var list=new List<string>();
    EnumWindows(delegate(IntPtr h, IntPtr l){
      uint p; GetWindowThreadProcessId(h, out p);
      if(p==pid){
        var c=new StringBuilder(256); GetClassName(h,c,256);
        var t=new StringBuilder(512); GetWindowText(h,t,512);
        RECT r; GetWindowRect(h, out r);
        list.Add(h.ToInt64() + " vis=" + (IsWindowVisible(h)?1:0) + " class=[" + c + "] title=[" + t + "] rect=" + r.L + "," + r.T + "," + r.R + "," + r.B);
      }
      return true; }, IntPtr.Zero);
    return list;
  }
  public static List<string> Kids(IntPtr parent, int depth){
    var list=new List<string>();
    if (depth > 6) return list;
    EnumChildWindows(parent, delegate(IntPtr h, IntPtr l){
      var c=new StringBuilder(256); GetClassName(h,c,256);
      var t=new StringBuilder(512); GetWindowText(h,t,512);
      RECT r; GetWindowRect(h, out r);
      list.Add("   kid vis=" + (IsWindowVisible(h)?1:0) + " class=[" + c + "] text=[" + t + "] rect=" + r.L + "," + r.T + "," + r.R + "," + r.B);
      return true; }, IntPtr.Zero);
    return list;
  }
}
'@
if ($TargetPid -eq 0) { $TargetPid = (Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -First 1).Id }
$lines = New-Object System.Collections.ArrayList
if (-not $TargetPid) { [void]$lines.Add('NO-POWERPNT') }
else {
  [void]$lines.Add('PID=' + $TargetPid)
  foreach ($l in [WL]::Top([uint32]$TargetPid)) { [void]$lines.Add($l) }
  if ($args.Count -gt 0 -and $args[0] -eq 'kids') {
    foreach ($l in @([WL]::Top([uint32]$TargetPid))) {
      if ($l -match '^(\-?\d+) vis=1') { $h = [IntPtr][int64]$Matches[1]; [void]$lines.Add('--- kids of ' + $h); foreach ($k in [WL]::Kids($h, 0)) { [void]$lines.Add($k) } }
    }
  }
}
[IO.File]::WriteAllLines((Join-Path $PSScriptRoot 'out\dbg-win.txt'), $lines, (New-Object Text.UTF8Encoding($true)))
Write-Host ('lines=' + $lines.Count)
