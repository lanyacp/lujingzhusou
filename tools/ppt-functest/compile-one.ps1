# Self-contained: open a file, force VBE full compile, watch for the error dialog.
param([string]$File = '', [int]$TimeoutSec = 60)
$ErrorActionPreference = 'Continue'
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class Cs2 {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr h);
  public static IntPtr FindDlg(uint pid){
    IntPtr f=IntPtr.Zero;
    EnumWindows(delegate(IntPtr h, IntPtr l){
      uint p; GetWindowThreadProcessId(h, out p);
      if(p==pid && IsWindowVisible(h)){ var c=new StringBuilder(64); GetClassName(h,c,64); if(c.ToString()=="#32770"){ f=h; return false; } }
      return true; }, IntPtr.Zero);
    return f;
  }
  public static string DlgText(IntPtr h){
    var txt=new StringBuilder(); var btns=new StringBuilder();
    EnumChildWindows(h, delegate(IntPtr c, IntPtr ll){
      var cls=new StringBuilder(64); GetClassName(c,cls,64);
      var t=new StringBuilder(1024); GetWindowText(c,t,1024);
      if(cls.ToString()=="Button"){ if(btns.Length>0) btns.Append("|"); btns.Append(t); }
      else if(t.Length>0){ if(txt.Length>0) txt.Append(" / "); txt.Append(t); }
      return true; }, IntPtr.Zero);
    return "text=[" + txt + "] buttons=[" + btns + "]";
  }
  public static void Ok(IntPtr dlg){
    EnumChildWindows(dlg, delegate(IntPtr c, IntPtr ll){
      var cls=new StringBuilder(64); GetClassName(c,cls,64);
      if(cls.ToString()=="Button"){ PostMessage(dlg, 0x0111, (IntPtr)GetDlgCtrlID(c), c); return false; }
      return true; }, IntPtr.Zero);
  }
}
'@
$log = New-Object System.Collections.ArrayList
$app = New-Object -ComObject PowerPoint.Application
$app.AutomationSecurity = 1
try { $app.Visible = -1 } catch {}
Start-Sleep -Milliseconds 2000
$ppid = (Get-Process POWERPNT | Select-Object -First 1).Id
if ($File) { $null = $app.Presentations.Open($File, -1, 0, -1); Start-Sleep -Milliseconds 1000 }
$vbe = $app.VBE
$proj = ''
try { $proj = $vbe.ActiveVBProject.Name } catch { $proj = '(none)' }
[void]$log.Add('activeProject=[' + $proj + ']')
# 找 调试菜单里 id=578 的 编译 命令
$ctl = $null
foreach ($c in $vbe.CommandBars.Item(1).Controls) {
  if ($c.Type -eq 10) { foreach ($cc in $c.Controls) { if ($cc.Id -eq 578) { $ctl = $cc } } }
}
if ($null -eq $ctl) { [void]$log.Add('compile cmd NOT FOUND') }
else {
  [void]$log.Add('compile cmd found, enabled=' + $ctl.Enabled)
  try { $ctl.Execute(); [void]$log.Add('execute() returned') } catch { [void]$log.Add('execute err: ' + $_.Exception.Message) }
}
$seen = @{}
$sw = [Diagnostics.Stopwatch]::StartNew()
while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
  $h = [Cs2]::FindDlg([uint32]$ppid)
  if ($h -ne [IntPtr]::Zero -and -not $seen.ContainsKey($h.ToInt64())) {
    $seen[$h.ToInt64()] = 1
    [void]$log.Add('DIALOG ' + [Cs2]::DlgText($h))
    [Cs2]::Ok($h)
    Start-Sleep -Milliseconds 800
    continue
  }
  # 编译命令执行完且没有弹窗 → 稳定即成功
  if ($log.Count -gt 0 -and (($log | Select-Object -Last 1) -match 'execute\(\) returned') -and $sw.Elapsed.TotalMilliseconds -gt 4000) { break }
  Start-Sleep -Milliseconds 200
}
$mode = ''
try { $mode = $vbe.Mode } catch {}
[void]$log.Add('vbeMode=' + $mode)
$errDialogs = @($log | Where-Object { $_ -match 'DIALOG' })
if ($errDialogs.Count -gt 0) { Write-Host 'RESULT: COMPILE-ERROR' } elseif ($mode -eq 0) { Write-Host 'RESULT: CLEAN' } else { Write-Host 'RESULT: UNKNOWN mode=' + $mode }
foreach ($l in $log) { Write-Host ('  ' + $l) }
try { foreach ($p in @($app.Presentations)) { $p.Saved = -1; $p.Close() }; $app.Quit() } catch {}
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
$sw2 = [Diagnostics.Stopwatch]::StartNew()
while ($sw2.Elapsed.TotalSeconds -lt 45 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue)) { Start-Sleep -Milliseconds 300 }
Write-Host ('exited=' + (-not [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)))
