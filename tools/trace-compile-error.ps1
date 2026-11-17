$ErrorActionPreference = 'Continue'
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 5
$tmp = Join-Path $env:TEMP ('abp_ce2_' + [guid]::NewGuid().ToString('N').Substring(0,8) + '.pptm')
Copy-Item 'D:\projects\wps插件\PowerPoint版\路径助手.pptm' $tmp -Force
$app = $null
for ($i=0; $i -lt 5 -and -not $app; $i++) { try { $app = New-Object -ComObject PowerPoint.Application } catch { Start-Sleep -Seconds 4 } }
if (-not $app) { Write-Host 'COM-FAIL'; exit 1 }
$app.AutomationSecurity = 1
try { $app.Visible = $true } catch {}
Start-Sleep -Milliseconds 1200
$pres = $app.Presentations.Open($tmp, -1, 0, -1)
Start-Sleep -Milliseconds 600
try { $app.VBE.MainWindow.Visible = $true } catch {}
Start-Sleep -Milliseconds 500
try { $app.VBE.ActiveVBProject = $pres.VBProject } catch {}
Start-Sleep -Milliseconds 300
function Find-Compile($controls) {
    foreach ($c in $controls) {
        if ($c.Caption -match 'Compile|编译') { $script:fc = $c; return $c.Id }
        try { if ($c.Type -eq 10) { $r = Find-Compile $c.Controls; if ($r) { return $r } } } catch {}
    }
    return 0
}
$cid = 0
foreach ($cb in $app.VBE.CommandBars) { $cid = Find-Compile $cb.Controls; if ($cid) { break } }
if ($cid -eq 0) { Write-Host 'NO-COMPILE-CTRL'; exit 1 }
Write-Host ('FOUND-COMPILE id=' + $cid + ' — EXECUTING (blocks on error dialog)')
$script:fc.Execute()
Write-Host 'COMPILE-RETURNED (no error!)'
