# READ-ONLY inspection of the user's live PowerPoint session. Never quits, never
# closes, never modifies: only reads and releases its own COM reference.
$log = New-Object System.Collections.ArrayList
$app = $null
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { [void]$log.Add('ATTACH FAILED: ' + $_.Exception.Message) }
if ($null -ne $app) {
  [void]$log.Add('version=' + $app.Version + ' windowState=' + $app.WindowState + ' presentations=' + $app.Presentations.Count)
  for ($i = 1; $i -le $app.Presentations.Count; $i++) {
    $p = $app.Presentations.Item($i)
    $fu = ''; try { $fu = $p.FullName } catch { $fu = '(unsaved)' }
    [void]$log.Add('PRES#' + $i + ' name=[' + $p.Name + '] path=[' + $fu + '] slides=' + $p.Slides.Count + ' saved=' + $p.Saved + ' windows=' + $p.Windows.Count + ' readOnly=' + $p.ReadOnly)
  }
  try {
    $aw = $app.ActiveWindow
    $fu = ''; try { $fu = $aw.Presentation.FullName } catch {}
    [void]$log.Add('ACTIVE window: pres=[' + $fu + '] viewType=' + $aw.ViewType + ' zoom=' + $aw.View.Zoom)
    try {
      $sl = $aw.View.Slide
      [void]$log.Add('  active slide index=' + $sl.SlideIndex + ' shapes=' + $sl.Shapes.Count)
    } catch { [void]$log.Add('  active slide: n/a ' + $_.Exception.Message) }
    $sel = $aw.Selection
    [void]$log.Add('  selection type=' + $sel.Type + ' (1=slides 2=shapes 3=text)')
    if ($sel.Type -eq 2) {
      $sr = $sel.ShapeRange
      for ($k = 1; $k -le $sr.Count; $k++) { $s = $sr.Item($k); [void]$log.Add('    selected: name=[' + $s.Name + '] id=' + $s.Id + ' type=' + $s.Type + ' L=' + [Math]::Round($s.Left) + ' T=' + [Math]::Round($s.Top)) }
    }
    # marker shapes anywhere in the active presentation
    $ap = $aw.Presentation
    for ($si = 1; $si -le $ap.Slides.Count; $si++) {
      $s2 = $ap.Slides.Item($si)
      for ($j = 1; $j -le $s2.Shapes.Count; $j++) {
        $sh = $s2.Shapes.Item($j)
        if ($sh.Name -like '__pthlp_*') { [void]$log.Add('  MARK slide ' + $si + ' name=[' + $sh.Name + '] id=' + $sh.Id + ' L=' + [Math]::Round($sh.Left, 1) + ' T=' + [Math]::Round($sh.Top, 1)) }
      }
      $ms = $s2.TimeLine.MainSequence.Count
      if ($ms -gt 0) {
        for ($e = 1; $e -le $ms; $e++) { $ef = $s2.TimeLine.MainSequence.Item($e); [void]$log.Add('  EFFECT slide ' + $si + ' #' + $e + ' shape=[' + $ef.Shape.Name + '] type=' + $ef.EffectType) }
      }
    }
  } catch { [void]$log.Add('ACTIVE window read err: ' + $_.Exception.Message) }
  try { foreach ($a in $app.AddIns) { [void]$log.Add('ADDIN name=[' + $a.Name + '] loaded=' + $a.Loaded + ' path=' + $a.Path) } } catch { [void]$log.Add('addins err: ' + $_.Exception.Message) }
  try {
    $vbe = $app.VBE
    [void]$log.Add('VBE mode=' + $vbe.Mode + ' projects=' + $vbe.VBProjects.Count)
    for ($k = 1; $k -le $vbe.VBProjects.Count; $k++) {
      $vp = $vbe.VBProjects.Item($k); $fn = ''
      try { $fn = $vp.FileName } catch {}
      $comps = @(); foreach ($c in $vp.VBComponents) { $comps += $c.Name }
      [void]$log.Add('  PROJECT name=[' + $vp.Name + '] file=[' + $fn + '] components=' + ($comps -join ','))
    }
  } catch { [void]$log.Add('vbe err: ' + $_.Exception.Message) }
  try { $cb = $app.CommandBars.Item('ABPathBar'); [void]$log.Add('ABPathBar visible=' + $cb.Visible + ' controls=' + $cb.Controls.Count) } catch { [void]$log.Add('ABPathBar missing') }
}
# visible top-level windows of the PowerPoint process (panels, dialogs)
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class LvWin {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  public static List<string> Vis(uint pid, string[] classes){
    var list=new List<string>();
    EnumWindows(delegate(IntPtr h, IntPtr l){
      uint p; GetWindowThreadProcessId(h, out p);
      if(p==pid && IsWindowVisible(h)){
        var c=new StringBuilder(256); GetClassName(h,c,256);
        var t=new StringBuilder(512); GetWindowText(h,t,512);
        string cs=c.ToString(); if(classes==null || Array.IndexOf(classes, cs)>=0) list.Add(cs + " | " + t.ToString());
      }
      return true; }, IntPtr.Zero);
    return list;
  }
}
'@
$pp = Get-Process POWERPNT -ErrorAction SilentlyContinue
[void]$log.Add('POWERPNT processes: ' + @($pp).Count)
foreach ($x in $pp) { [void]$log.Add('  pid=' + $x.Id + ' start=' + $x.StartTime + ' title=[' + $x.MainWindowTitle + ']') }
if ($pp) { try { foreach ($w in [LvWin]::Vis([uint32]($pp | Select-Object -First 1).Id, @('PPTFrameClass','ThunderDFrame','#32770','wndclass_desked_gsk'))) { [void]$log.Add('  WIN ' + $w) } } catch { [void]$log.Add('  win enum err: ' + $_.Exception.Message) } }
if ($null -ne $app) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {} }
[IO.File]::WriteAllLines((Join-Path $PSScriptRoot 'out\inspect-live.txt'), $log, (New-Object Text.UTF8Encoding($true)))
Write-Host ('lines=' + $log.Count)
