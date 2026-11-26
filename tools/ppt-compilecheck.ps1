<#
  ppt-compilecheck.ps1 - force a FULL compile of a VBA project inside the running
  PowerPoint, through the VBE's own "Compile" command (menu id 578).

  Why: VBA compiles ON DEMAND. A project can load and run fine while still hiding
  compile errors in procedures nobody has called yet - exactly how a "compile error in
  a hidden module" pops up only when the user first clicks a particular button.
  Running the full compile at build time surfaces every such error up front.

  A dialog watchdog captures and dismisses the compile-error dialog if one appears.
  Exit codes: 0 = clean, 1 = compile error, 2 = could not run.
  Usage: powershell -File tools\ppt-compilecheck.ps1 [-Match <project file/name fragment>]
#>
param([string]$Match = '', [int]$TimeoutSec = 90)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class CsSpy {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr h);
  public class Dlg { public IntPtr H; public string Title=""; public string Text=""; public List<IntPtr> Hs = new List<IntPtr>(); }
  static string T(IntPtr h){ var sb=new StringBuilder(2048); GetWindowText(h,sb,2048); return sb.ToString(); }
  static string C(IntPtr h){ var sb=new StringBuilder(256); GetClassName(h,sb,256); return sb.ToString(); }
  public static List<Dlg> Scan(uint pid){
    var list=new List<Dlg>();
    EnumWindows(delegate(IntPtr h, IntPtr l){
      uint p; GetWindowThreadProcessId(h, out p);
      if(p==pid && IsWindowVisible(h) && C(h)=="#32770"){
        Dlg d=new Dlg(); d.H=h; d.Title=T(h);
        StringBuilder txt=new StringBuilder();
        EnumChildWindows(h, delegate(IntPtr c, IntPtr ll){
          string cls=C(c), t=T(c);
          if(cls=="Button") d.Hs.Add(c);
          else if(t.Length>0){ if(txt.Length>0) txt.Append(" / "); txt.Append(t); }
          return true; }, IntPtr.Zero);
        d.Text=txt.ToString(); list.Add(d);
      }
      return true; }, IntPtr.Zero);
    return list;
  }
  public static void Ok(IntPtr dlg){     // click the first button (OK on a compile-error box)
    EnumChildWindows(dlg, delegate(IntPtr c, IntPtr ll){
      if(C(c)=="Button"){ PostMessage(dlg, 0x0111, (IntPtr)GetDlgCtrlID(c), c); return false; }
      return true; }, IntPtr.Zero);
  }
}
'@

function Find-CompileCtl($vbe) {
  $ctl = $null
  try {
    foreach ($c in $vbe.CommandBars.Item(1).Controls) {          # VBE menu bar
      if ($c.Type -eq 10) {
        foreach ($cc in $c.Controls) { if ($cc.Id -eq 578) { $ctl = $cc } }
      }
    }
  } catch {}
  return $ctl
}

$app = $null
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch {}
if ($null -eq $app) { Write-Host 'COMPILE-CHECK: NO POWERPOINT'; exit 2 }
$script:ppid = (Get-Process POWERPNT | Select-Object -First 1).Id
$vbe = $app.VBE
$ctl0 = Find-CompileCtl $vbe
if ($null -eq $ctl0) { Write-Host 'COMPILE-CHECK: compile command not found'; exit 2 }

$job = Start-Job -ArgumentList $Match -ScriptBlock {
  param($Match)
  try {
    $a = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    $v = $a.VBE
    if ($Match) {
      for ($k = 1; $k -le $v.VBProjects.Count; $k++) {
        $p = $v.VBProjects.Item($k); $fn = ''
        try { $fn = $p.FileName } catch {}
        if (($fn + $p.Name) -like ('*' + $Match + '*')) { $v.ActiveVBProject = $p; break }
      }
    }
    $ctl = $null
    foreach ($c in $v.CommandBars.Item(1).Controls) {
      if ($c.Type -eq 10) { foreach ($cc in $c.Controls) { if ($cc.Id -eq 578) { $ctl = $cc } } }
    }
    if ($null -eq $ctl) { 'NOCTL'; return }
    $ctl.Execute()
    'EXECUTED'
  } catch { 'JOBERR ' + $_.Exception.Message }
}

$script:seen = @{}
$script:dialogs = New-Object System.Collections.ArrayList
function Watch-Once {
  foreach ($d in [CsSpy]::Scan([uint32]$script:ppid)) {
    if (-not $script:seen.ContainsKey($d.H.ToInt64())) {
      $script:seen[$d.H.ToInt64()] = 1
      [void]$script:dialogs.Add('title=[' + $d.Title + '] text=[' + ($d.Text -replace "`r`n", ' ') + ']')
      if ($d.Hs.Count -gt 0) { [CsSpy]::Ok($d.H) } else { [void][CsSpy]::PostMessage($d.H, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) }
    }
  }
}
$sw = [Diagnostics.Stopwatch]::StartNew(); $timedOut = $false
while ($job.State -eq 'Running') {
  Watch-Once
  if ($sw.Elapsed.TotalSeconds -gt $TimeoutSec) { $timedOut = $true; break }
  Start-Sleep -Milliseconds 120
}
$jobOut = ''
if (-not $timedOut) { $jobOut = ((Receive-Job $job | Out-String)).Trim() }
Remove-Job $job -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 900
Watch-Once

$mode = ''
try { $mode = $vbe.Mode } catch {}
$en = ''
try { $c1 = Find-CompileCtl $vbe; $en = $c1.Enabled } catch {}
Write-Host ('COMPILE-CHECK: job=[' + $jobOut + '] vbeMode=' + $mode + ' compileCmdEnabled=' + $en)
$errs = @($script:dialogs | Where-Object { $_ -match 'Visual Basic' })
foreach ($d in $script:dialogs) { Write-Host ('COMPILE-CHECK: dialog ' + $d) }
if ($timedOut) { Write-Host 'COMPILE-CHECK: FAIL (timed out)'; exit 1 }
if ($errs.Count -gt 0) { Write-Host 'COMPILE-CHECK: FAIL (compile error)'; exit 1 }
if ($jobOut -notlike 'EXECUTED*') { Write-Host 'COMPILE-CHECK: WARN (compile command did not run)'; exit 2 }
if ($en) { Write-Host 'COMPILE-CHECK: PASS (no error dialog; command still enabled)'; exit 0 }
Write-Host 'COMPILE-CHECK: PASS (whole project compiled; command greyed out)'
exit 0
