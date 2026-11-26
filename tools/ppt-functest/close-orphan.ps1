# Close an orphaned PowerPoint test instance gracefully: WM_CLOSE via taskkill (no /F),
# auto-answering the "save changes?" prompt with Don't Save. Never force-kills.
param([int]$TargetPid = 0)
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
if ($TargetPid -eq 0) { $p = Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -First 1; if ($p) { $TargetPid = $p.Id } }
if ($TargetPid -eq 0 -or -not (Get-Process -Id $TargetPid -ErrorAction SilentlyContinue)) { Write-Host 'NOT-RUNNING'; exit 0 }
Write-Host ('target pid=' + $TargetPid)
$out = & taskkill.exe /PID $TargetPid 2>&1
Write-Host ('taskkill(graceful): ' + (($out | Out-String).Trim()))
$rx = [string][char]0x4E0D + [string][char]0x4FDD + [string][char]0x5B58          # 不保存
$condPid = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, [int]$TargetPid)
$sw = [Diagnostics.Stopwatch]::StartNew()
while ($sw.Elapsed.TotalSeconds -lt 40 -and (Get-Process -Id $TargetPid -ErrorAction SilentlyContinue)) {
  try {
    $wins = [System.Windows.Automation.AutomationElement]::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $condPid)
    foreach ($w in $wins) {
      $bc = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
      foreach ($b in $w.FindAll([System.Windows.Automation.TreeScope]::Descendants, $bc)) {
        $n = $b.Current.Name
        if ($n -match $rx -or $n -match "Don.t Save") {
          try { $b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke(); Write-Host ('clicked [' + $n + ']') } catch { Write-Host ('invoke err: ' + $_.Exception.Message) }
        }
      }
    }
  } catch {}
  Start-Sleep -Milliseconds 400
}
$alive = [bool](Get-Process -Id $TargetPid -ErrorAction SilentlyContinue)
Write-Host ('exited=' + (-not $alive))
if (-not $alive) { exit 0 } else { exit 1 }
