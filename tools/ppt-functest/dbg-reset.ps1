Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Runtime.InteropServices;
public static class WMove {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
  [DllImport("user32.dll")] public static extern bool AllowSetForegroundWindow(int pid);
  public static IntPtr Find(uint pid, string cls){
    IntPtr f=IntPtr.Zero;
    EnumWindows(delegate(IntPtr h, IntPtr l){
      uint p; GetWindowThreadProcessId(h, out p);
      if(p==pid){ var c=new StringBuilder(256); GetClassName(h,c,256); if(c.ToString()==cls){ f=h; return false; } }
      return true; }, IntPtr.Zero);
    return f;
  }
  public static string Title(IntPtr h){ var sb=new StringBuilder(512); GetWindowText(h,sb,512); return sb.ToString(); }
  public static string OnScreen(IntPtr h){
    ShowWindow(h, 9);                                   // SW_RESTORE
    SetWindowPos(h, IntPtr.Zero, 80, 80, 1100, 720, 0x40);   // SWP_SHOWWINDOW, HWND_TOP
    BringWindowToTop(h);
    bool fg = SetForegroundWindow(h);
    return "SetForegroundWindow=" + fg + " isFg=" + (GetForegroundWindow()==h);
  }
}
'@
$log = New-Object System.Collections.ArrayList
$rx = 'Reset|' + [char]0x91CD + [char]0x7F6E          # Reset | 重置
$pp = Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $pp) { [void]$log.Add('NO-POWERPNT') }
else {
  $h = [WMove]::Find([uint32]$pp.Id, 'wndclass_desked_gsk')
  [void]$log.Add('VBE hwnd=' + $h + ' title=[' + [WMove]::Title($h) + ']')
  [void]$log.Add([WMove]::OnScreen($h))
  Start-Sleep -Milliseconds 800
  if ([WMove]::GetForegroundWindow() -ne $h) {
    try { $wsh = New-Object -ComObject WScript.Shell; [void]$log.Add('AppActivate(pid)=' + $wsh.AppActivate($pp.Id)) } catch { [void]$log.Add('AppActivate err') }
    Start-Sleep -Milliseconds 600
  }
  [void]$log.Add('foreground is VBE: ' + ([WMove]::GetForegroundWindow() -eq $h))
  $invoked = $false
  try {
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($h)
    $all = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
    [void]$log.Add('UIA descendants=' + $all.Count)
    $cands = @()
    foreach ($e in $all) {
      $n = ''; $t = ''
      try { $n = $e.Current.Name; $t = $e.Current.ControlType.ProgrammaticName } catch {}
      if ($n -match $rx) { $cands += , @($n, $t, $e) }
    }
    foreach ($c in $cands) { [void]$log.Add('  cand [' + $c[1] + '] name=[' + $c[0] + ']') }
    foreach ($c in $cands) {
      if ($c[1] -ne 'ControlType.Button') { continue }
      try { $p = $c[2].GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern); $p.Invoke(); $invoked = $true; [void]$log.Add('  invoked [' + $c[0] + ']'); break }
      catch { [void]$log.Add('  invoke err: ' + $_.Exception.Message) }
    }
  } catch { [void]$log.Add('UIA ERR: ' + $_.Exception.Message) }
  if (-not $invoked -and [WMove]::GetForegroundWindow() -eq $h) {
    $wsh = New-Object -ComObject WScript.Shell
    $wsh.SendKeys('%r'); Start-Sleep -Milliseconds 600
    $wsh.SendKeys('{DOWN}{DOWN}{ENTER}'); Start-Sleep -Milliseconds 1500
    [void]$log.Add('sent Alt+R, Down, Down, Enter')
  }
  Start-Sleep -Milliseconds 1500
  [void]$log.Add('VBE title now=[' + [WMove]::Title($h) + ']')
  try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    [void]$log.Add('presentations=' + $app.Presentations.Count)
    $app.Quit()
    [void]$log.Add('Quit OK')
  } catch { [void]$log.Add('COM ERR: ' + $_.Exception.Message) }
}
$sw = [Diagnostics.Stopwatch]::StartNew()
while ($sw.Elapsed.TotalSeconds -lt 25 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue)) { Start-Sleep -Milliseconds 300 }
[void]$log.Add('POWERPNT alive after wait: ' + [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue))
[IO.File]::WriteAllLines((Join-Path $PSScriptRoot 'out\dbg-reset.txt'), $log, (New-Object Text.UTF8Encoding($true)))
