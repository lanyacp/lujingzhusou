# 让(测试拉起的) PowerPoint 干净退出: 有用户文稿则不动, 除非 -Force
param([switch]$Force)
$log = New-Object System.Collections.ArrayList
$pp = Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $pp) { [void]$log.Add('no POWERPNT') }
else {
  try {
    $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    $n = $app.Presentations.Count
    [void]$log.Add('presentations=' + $n)
    if ($n -eq 0 -or $Force) {
      if ($Force) { for ($i = 0; $i -lt $n; $i++) { $p = $app.Presentations.Item(1); [void]$log.Add('closing ' + $p.Name); $p.Saved = -1; $p.Close() } }
      try { $app.Quit(); [void]$log.Add('Quit OK') } catch { [void]$log.Add('Quit ERR ' + $_.Exception.Message) }
    } else { [void]$log.Add('user presentations open -> left alone') }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
  } catch { [void]$log.Add('COM ERR ' + $_.Exception.Message) }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  $sw = [Diagnostics.Stopwatch]::StartNew()
  while ($sw.Elapsed.TotalSeconds -lt 60 -and (Get-Process POWERPNT -ErrorAction SilentlyContinue)) { Start-Sleep -Milliseconds 300 }
  [void]$log.Add('exited=' + (-not [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)))
}
[IO.File]::WriteAllLines((Join-Path $PSScriptRoot 'out\quit-pp.txt'), $log, (New-Object Text.UTF8Encoding($true)))
foreach ($l in $log) { Write-Host $l }
