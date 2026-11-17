$ErrorActionPreference = 'Continue'
Start-Sleep -Seconds 2
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { Write-Host 'B-ATTACH-FAIL'; exit 1 }
$out = @()
try {
    $cp = $app.VBE.ActiveCodePane
    if ($cp) {
        $sl=0;$sc=0;$el=0;$ec=0
        $cp.GetSelection([ref]$sl,[ref]$sc,[ref]$el,[ref]$ec)
        $out += ('ERROR-LINE-NUM=' + $sl)
        $out += ('ERROR-MODULE=' + $cp.CodeModule.Parent.Name)
        $lo = [Math]::Max(1,$sl-3); $hi=[Math]::Min($cp.CodeModule.CountOfLines,$sl+3)
        for ($i=$lo;$i -le $hi;$i++) { $out += ('L' + $i + '|' + $cp.CodeModule.Lines($i,1)) }
    } else { $out += 'NO-ACTIVE-PANE' }
} catch { $out += ('B-READ-ERR: ' + $_.Exception.Message) }
[IO.File]::WriteAllLines('D:\projects\wps插件\tools\error-line.txt', $out, [System.Text.Encoding]::UTF8)
Write-Host 'B-WROTE'
# 关闭模态对话框(编译错误 MsgBox)
try {
    $wsh = New-Object -ComObject WScript.Shell
    $wsh.AppActivate('Microsoft Visual Basic') | Out-Null
    Start-Sleep -Milliseconds 400
    $wsh.SendKeys('{ENTER}')
    Start-Sleep -Milliseconds 400
    $wsh.SendKeys('{ENTER}')
} catch {}
Write-Host 'B-DONE'
