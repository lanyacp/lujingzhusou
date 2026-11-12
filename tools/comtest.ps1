# comtest.ps1 - FINAL end-to-end COM verification, mirrors js/core.js production logic
# NOTE: Sequence COM objects get enumerated when returned from PS functions -> always inline
#       "$slide.TimeLine.MainSequence" at each use site (fresh-refetch discipline).
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File tools\comtest.ps1
$ErrorActionPreference = 'Continue'

$pass = 0; $fail = 0
function Check($cond, $msg) {
    if ($cond) { $script:pass++; Write-Host "  [OK]  $msg" }
    else { $script:fail++; Write-Host "  [FAIL] $msg" }
}
# find motion behavior of an Effect, store into $script:motion (never return COM collections)
function FindMotion($eff) {
    $script:motion = $null
    try {
        $bc = $eff.Behaviors.Count
        for ($i = 1; $i -le $bc; $i++) {
            if ($eff.Behaviors.Item($i).Type -eq 1) { $script:motion = $eff.Behaviors.Item($i).MotionEffect; return }
        }
    } catch { }
}

$tmpDir = Join-Path $env:TEMP "abp_comtest"
if (Test-Path $tmpDir) { Remove-Item -Recurse -Force $tmpDir }
New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
$pptx = Join-Path $tmpDir "comtest.pptx"

$app = New-Object -ComObject KWPP.Application
try { $app.Visible = $true } catch { }
Start-Sleep -Milliseconds 600
$pres = $app.Presentations.Add(-1)
try {

    $slideW = $pres.PageSetup.SlideWidth
    $slideH = $pres.PageSetup.SlideHeight
    $slide = $pres.Slides.Add(1, 12)

    # --- shapes: mover rect center(160,140); target oval center(540,340) ---
    $rect = $slide.Shapes.AddShape(1, 100, 100, 120, 80)
    $oval = $slide.Shapes.AddShape(9, 500, 300, 80, 80)
    $ax = $rect.Left + $rect.Width / 2;  $ay = $rect.Top + $rect.Height / 2
    $bx = $oval.Left + $oval.Width / 2;  $by = $oval.Top + $oval.Height / 2
    $dx = ($bx - $ax) / $slideW;  $dy = ($by - $ay) / $slideH
    Write-Host ("  A=($ax,$ay) B=($bx,$by) dx={0:F6} dy={1:F6}" -f $dx, $dy)

    # --- marker shapes like core.js addCross / preview line ---
    $aMark = $null; $bMark = $null; $line = $null
    try { $aMark = $slide.Shapes.AddShape(11, $ax - 7, $ay - 7, 14, 14); $aMark.Name = '__pthlp_A__' } catch { Write-Host ("  A mark err: " + $_.Exception.Message) }
    try { $bMark = $slide.Shapes.AddShape(11, $bx - 8, $by - 8, 16, 16); $bMark.Name = '__pthlp_B__' } catch { Write-Host ("  B mark err: " + $_.Exception.Message) }
    try {
        $line = $slide.Shapes.AddLine($ax, $ay, $bx, $by)
        $line.Name = '__pthlp_LINE__'
        $line.Line.DashStyle = 4; $line.Line.Weight = 1.25
    } catch { Write-Host ("  line err: " + $_.Exception.Message) }
    Check ($null -ne $aMark -and $null -ne $bMark -and $null -ne $line) "markers + preview line created"

    $bmcX = $bMark.Left + $bMark.Width / 2; $bmcY = $bMark.Top + $bMark.Height / 2
    Check (([math]::Abs($bmcX - $bx) -lt 0.01) -and ([math]::Abs($bmcY - $by) -lt 0.01)) "B marker center == target shape center"

    # --- core.js createPathEffect: Strategy 1 (all inline, fresh each time) ---
    $pathStr = "M 0 0 L {0:F6} {1:F6} E" -f $dx, $dy
    $eff = $slide.TimeLine.MainSequence.AddEffect($rect, 149, 0, 1)
    Check ($null -ne $eff) "S1: AddEffect(preset path) ok"
    FindMotion $eff
    Check ($null -ne $script:motion) "S1: motion behavior found"
    $script:motion.Path = $pathStr
    # fresh re-walk readback
    $c1 = $slide.TimeLine.MainSequence.Count
    $e1 = $slide.TimeLine.MainSequence.Item($c1)
    FindMotion $e1
    $rb = [string]$script:motion.Path
    Write-Host "  S1 readback(fresh) = '$rb'"
    Check ($rb -match ([regex]::Escape($dx.ToString('F6'))) -and $rb -match ([regex]::Escape($dy.ToString('F6')))) "S1: Path write + fresh readback match"

    # --- Timing via fresh last effect ---
    $c2 = $slide.TimeLine.MainSequence.Count
    $te = $slide.TimeLine.MainSequence.Item($c2)
    $te.Timing.Duration = 2.0
    $te.Timing.SmoothStart = 0; $te.Timing.SmoothEnd = 0
    $te.Timing.TriggerType = 1
    $c3 = $slide.TimeLine.MainSequence.Count
    $durBack = $slide.TimeLine.MainSequence.Item($c3).Timing.Duration
    Check ([math]::Abs($durBack - 2.0) -lt 0.01) "Timing(Duration/Smooth/Trigger) via fresh ref"

    # --- replaceAnim: add 2nd effect then deleteShapeAnimations loop ---
    $null = $slide.TimeLine.MainSequence.AddEffect($rect, 149, 0, 1)
    Check ($slide.TimeLine.MainSequence.Count -eq 2) "second effect added (count=2)"
    $deleted = 0; $guard = 50
    $targetName = $rect.Name
    $i2 = $slide.TimeLine.MainSequence.Count
    while ($i2 -ge 1 -and $guard-- -gt 0) {
        $hit = $false
        try {
            $e2 = $slide.TimeLine.MainSequence.Item($i2)
            if ($e2.Shape.Name -eq $targetName) { $hit = $true }
        } catch { }
        if ($hit) {
            try { $slide.TimeLine.MainSequence.Item($i2).Delete(); $deleted++ } catch { }
            $i2 = $slide.TimeLine.MainSequence.Count
        } else { $i2-- }
    }
    Check ($deleted -eq 2 -and $slide.TimeLine.MainSequence.Count -eq 0) "deleteShapeAnimations loop removed all 2 effects"

    # --- regenerate final effect for XML ground truth (markers still on slide) ---
    $effF = $slide.TimeLine.MainSequence.AddEffect($rect, 149, 0, 1)
    FindMotion $effF
    $script:motion.Path = $pathStr
    $c4 = $slide.TimeLine.MainSequence.Count
    $teF = $slide.TimeLine.MainSequence.Item($c4)
    $teF.Timing.Duration = 2.0; $teF.Timing.TriggerType = 1

    # --- cleanFlowMarkers equivalent ---
    $victims = @()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $nm = $slide.Shapes.Item($i).Name
        if ($nm -like '__pthlp_*') { $victims += $i }
    }
    for ($j = $victims.Count - 1; $j -ge 0; $j--) { $slide.Shapes.Item($victims[$j]).Delete() }
    Check ($victims.Count -eq 3) "cleanFlowMarkers removed 3 helper shapes"

    # --- save & XML ground truth ---
    $pres.SaveAs($pptx, 24)
    Check (Test-Path $pptx) "pptx saved"
    Copy-Item $pptx (Join-Path $tmpDir 'copy.zip') -Force
    Expand-Archive (Join-Path $tmpDir 'copy.zip') (Join-Path $tmpDir 'x') -Force
    $sxp = (Get-ChildItem (Join-Path $tmpDir 'x\ppt\slides') -Filter 'slide*.xml' | Select-Object -First 1).FullName
    $xml = [IO.File]::ReadAllText($sxp)
    $m = [regex]::Match($xml, '<p:animMotion[^>]*\bpath="([^"]+)"')
    if ($m.Success) {
        $xp = $m.Groups[1].Value
        $lm = [regex]::Match($xp, 'L\s*(-?[\d.eE+-]+)\s+(-?[\d.eE+-]+)')
        $xdx = [double]$lm.Groups[1].Value; $xdy = [double]$lm.Groups[2].Value
        Write-Host ("  XML animMotion path = `"$xp`"")
        Check (([math]::Abs($xdx - $dx) -lt 1e-4) -and ([math]::Abs($xdy - $dy) -lt 1e-4)) "XML ground truth: mover center -> target center displacement exact"
        Check (-not ($xml -match '__pthlp_')) "helper shapes cleaned in saved file"
    } else {
        Check $false "animMotion path found in slide XML"
    }

}
finally {
    try { $pres.Close() } catch { }
    try { $app.Quit() } catch { }
    Start-Sleep -Milliseconds 800
}

Write-Host ""
Write-Host ("RESULT: pass=$pass fail=$fail")
Remove-Item -Recurse -Force $tmpDir -ErrorAction SilentlyContinue
exit ($(if ($fail -eq 0) { 0 } else { 1 }))
