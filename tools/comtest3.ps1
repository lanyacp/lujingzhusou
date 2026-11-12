# comtest3.ps1 - verify "stale COM proxy" theory: always re-fetch TimeLine.MainSequence after mutations
$ErrorActionPreference = 'Continue'
function Step($name, $fn) {
    try { $r = & $fn; if ($null -eq $r) { $r = '' }; Write-Host ("  [OK]  $name $r") }
    catch { Write-Host ("  [ERR] $name :: " + $_.Exception.Message) }
}

$app = New-Object -ComObject KWPP.Application
try { $app.Visible = $true } catch { }
Start-Sleep -Milliseconds 600
$pres = $app.Presentations.Add(-1)
$slide = $pres.Slides.Add(1, 12)
$rect = $slide.Shapes.AddShape(1, 100, 100, 120, 80)

Write-Host "== 1: fresh-refetch discipline =="
Step 'AddEffect (fresh seq)' { $s = $slide.TimeLine.MainSequence; $s.AddEffect($rect, 149, 0, 1) | Out-Null; 'added' }
Step 'fresh count after add' { $slide.TimeLine.MainSequence.Count }
Step 'stale-count control (same-line)' { $s2 = $slide.TimeLine.MainSequence; $s2.Count }
Step 'fetch fresh Item(1).Behaviors.Count' { $slide.TimeLine.MainSequence.Item(1).Behaviors.Count }

Write-Host "== 2: Path write on fresh refs =="
Step 'write Path fresh' {
    $me = $slide.TimeLine.MainSequence.Item(1).Behaviors.Item(1).MotionEffect
    $me.Path = 'M 0 0 L 0.312500 0.203704 E'
    'wrote'
}
Step 'read Path fresh' { [string]$slide.TimeLine.MainSequence.Item(1).Behaviors.Item(1).MotionEffect.Path }

Write-Host "== 3: FromX/ToX on fresh refs (retest after Path write) =="
Step 'write From/To fresh' {
    $me = $slide.TimeLine.MainSequence.Item(1).Behaviors.Item(1).MotionEffect
    $me.FromX = 0; $me.FromY = 0; $me.ToX = 0.42; $me.ToY = 0.17
    'wrote'
}
Step 'read ToX/ToY fresh' {
    $me = $slide.TimeLine.MainSequence.Item(1).Behaviors.Item(1).MotionEffect
    "ToX=$($me.ToX) ToY=$($me.ToY)"
}
Step 'read Path after From/To' { [string]$slide.TimeLine.MainSequence.Item(1).Behaviors.Item(1).MotionEffect.Path }

Write-Host "== 4: Effect.Delete with fresh refs =="
Step 'delete via fresh Item(1)' {
    $slide.TimeLine.MainSequence.Item(1).Delete()
    'deleted'
}
Step 'fresh count after delete' { $slide.TimeLine.MainSequence.Count }
Step 'app alive' { $app.Presentations.Count }

Write-Host "== 5: Custom + Behaviors.Add + Path (fresh refs) =="
Step 'AddEffect custom' { $slide.TimeLine.MainSequence.AddEffect($rect, 0) | Out-Null; 'added' }
Step 'fresh count' { $slide.TimeLine.MainSequence.Count }
Step 'Behaviors.Add(motion)' {
    $effC = $slide.TimeLine.MainSequence.Item(1)
    $script:beh = $effC.Behaviors.Add(1)
    "behaviors=$($slide.TimeLine.MainSequence.Item(1).Behaviors.Count)"
}
Step 'write Path on added behavior' { $script:beh.MotionEffect.Path = 'M 0 0 L 0.111111 0.222222 E'; 'wrote' }
Step 'read Path on added behavior (fresh)' {
    [string]$slide.TimeLine.MainSequence.Item(1).Behaviors.Item(1).MotionEffect.Path
}

Write-Host "== 6: delete-all loop with fresh refs =="
Step 'delete all' {
    $n = $slide.TimeLine.MainSequence.Count
    for ($i = $n; $i -ge 1; $i--) { $slide.TimeLine.MainSequence.Item($i).Delete() }
    "before=$n after=$($slide.TimeLine.MainSequence.Count)"
}
Step 'app alive at end' { $app.Presentations.Count }

try { $pres.Close() } catch { Write-Host ("close: " + $_.Exception.Message) }
try { $app.Quit() } catch { }
Write-Host 'done'
