# Read-only probe: which tabs does the PowerPoint ribbon actually show?
# Uses an already running instance if it only holds our test presentation,
# otherwise starts a fresh one. Never closes user documents.
param([switch]$Fresh)
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$log = New-Object System.Collections.ArrayList
$app = $null
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch {}
$userDocs = $false
if ($null -ne $app) {
  try {
    $n = $app.Presentations.Count
    [void]$log.Add('existing instance: presentations=' + $n)
    for ($i = 1; $i -le $n; $i++) {
      $p = $app.Presentations.Item($i)
      $nm = $p.Name; $fu = ''
      try { $fu = $p.FullName } catch {}
      [void]$log.Add('  pres: ' + $nm + ' path=[' + $fu + '] saved=' + $p.Saved)
      if ($fu -ne '' -or ($nm -notlike '*1' -and $nm -notlike 'ABP_*')) { $userDocs = $true }
    }
  } catch { $userDocs = $true; [void]$log.Add('  could not list presentations -> treat as user instance') }
}
if ($null -eq $app) {
  $app = New-Object -ComObject PowerPoint.Application
  try { $app.Visible = -1 } catch {}
  [void]$log.Add('started new instance')
}
if (-not $userDocs) {
  try { $app.WindowState = 3 } catch {}
  if ($app.Presentations.Count -eq 0) { $null = $app.Presentations.Add(-1) }
  Start-Sleep -Milliseconds 2500
  $ppid = (Get-Process POWERPNT | Select-Object -First 1).Id
  $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, [int]$ppid)
  $wins = [System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)
  [void]$log.Add('top-level windows of PowerPoint: ' + $wins.Count)
  foreach ($w in $wins) {
    [void]$log.Add(' window class=[' + $w.Current.ClassName + '] name=[' + $w.Current.Name + ']')
    $tc = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::TabItem)
    $tabs = $w.FindAll([System.Windows.Automation.TreeScope]::Descendants, $tc)
    if ($tabs.Count -gt 0) {
      $names = @(); foreach ($t in $tabs) { $names += $t.Current.Name }
      [void]$log.Add('   TABS: ' + ($names -join ' | '))
    }
  }
  # add-in side facts
  try {
    foreach ($a in $app.AddIns) { [void]$log.Add('ADDIN ' + $a.Name + ' loaded=' + $a.Loaded + ' registered=' + $a.Registered + ' path=' + $a.Path) }
    $cb = $app.CommandBars.Item('ABPathBar'); [void]$log.Add('ABPathBar visible=' + $cb.Visible + ' controls=' + $cb.Controls.Count + ' pos=' + $cb.Position)
  } catch { [void]$log.Add('add-in check err: ' + $_.Exception.Message) }
}
if ($null -ne $app) {
  if (-not $userDocs) { try { $app.Quit() } catch {} }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
$sw = [Diagnostics.Stopwatch]::StartNew()
while ($sw.Elapsed.TotalSeconds -lt 45 -and -not $userDocs -and (Get-Process POWERPNT -ErrorAction SilentlyContinue)) { Start-Sleep -Milliseconds 300 }
[void]$log.Add('exited=' + (-not [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)))
[IO.File]::WriteAllLines((Join-Path $PSScriptRoot 'out\dbg-ribbon.txt'), $log, (New-Object Text.UTF8Encoding($true)))
foreach ($l in $log) { Write-Host $l }
