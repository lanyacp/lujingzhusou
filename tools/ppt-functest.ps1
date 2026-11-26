<#
  ppt-functest.ps1  —  路径助手 PowerPoint 版 功能测试编排器
  作用: 起一个临时演示文稿, 注入 VBA 测试驱动(tools/ppt-functest/*.bas),
        调用其入口宏, 期间用看门狗记录并关闭 PowerPoint 弹出的对话框
        (VBA 运行时错误 / MsgBox / 编译错误), 最后把结果落盘并打印.
  用法: powershell -File tools\ppt-functest.ps1 -Driver <bas> -Entry <宏名> [-Policy ok|yes|no|cancel]
  注意: 本文件含中文, 必须保持 UTF-8 BOM (Windows PowerShell 5.1 无 BOM 会按 GBK 误读)
#>
param(
  [string]$Driver = (Join-Path $PSScriptRoot 'ppt-functest\ft-driver.bas'),
  [string]$Entry  = 'FT_All',
  [string]$Policy = 'ok',
  [string]$Tag    = 'ft',
  [int]$TimeoutSec = 300,
  [string[]]$OpenFirst = @(),     # 先于临时文稿打开的文件(如 路径助手.pptm)
  [switch]$NoScratch,             # 不建临时文稿(驱动注入 OpenFirst 之后的活动文稿)
  [switch]$NoDismiss,             # 只记录弹窗不关闭(调试用)
  [string]$ErrorBtn = 'end',       # VBA 运行时错误框点哪个按钮: end(结束) | debug(调试, 并读取出错行) | continue
  [string]$ClickTitle = '',        # 标题匹配该正则的窗口出现后, 模拟真人鼠标点一下它的左上内侧(激活窗体)
  [string]$SimpleBas = '',         # 额外注入的标准模块(如 路径助手-简易宏.bas), 模块名取自其 Attribute VB_Name
  [string]$ShotOnDlg = '',         # 弹出 VBA 错误框时, 先把 VBE 截图存到 out\<该值>.png 再关闭
  [switch]$HoldOnBreak,            # 点了"调试"后停在现场, 直到出现 %TEMP%\abp_ft_go (便于外部排查)
  [switch]$KeepOpen               # 结束后不退出 PowerPoint
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class DlgSpy {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
  public class Dlg { public IntPtr H; public string Title=""; public string Text=""; public string Buttons=""; public List<IntPtr> Hs = new List<IntPtr>(); }
  static string T(IntPtr h){ var sb=new StringBuilder(2048); GetWindowText(h,sb,2048); return sb.ToString(); }
  static string C(IntPtr h){ var sb=new StringBuilder(256); GetClassName(h,sb,256); return sb.ToString(); }
  public static List<Dlg> Scan(uint pid){
    var list=new List<Dlg>();
    EnumWindows(delegate(IntPtr h, IntPtr l){
      uint p; GetWindowThreadProcessId(h, out p);
      if(p==pid && IsWindowVisible(h) && C(h)=="#32770"){
        Dlg d=new Dlg(); d.H=h; d.Title=T(h);
        StringBuilder txt=new StringBuilder(); StringBuilder btns=new StringBuilder();
        EnumChildWindows(h, delegate(IntPtr c, IntPtr ll){
          string cls=C(c), t=T(c);
          if(cls=="Button"){ if(btns.Length>0) btns.Append("|"); btns.Append(t); d.Hs.Add(c); }
          else if(t.Length>0){ if(txt.Length>0) txt.Append(" / "); txt.Append(t); }
          return true; }, IntPtr.Zero);
        d.Text=txt.ToString(); d.Buttons=btns.ToString(); list.Add(d);
      }
      return true; }, IntPtr.Zero);
    return list;
  }
  public static void Click(IntPtr dlg, IntPtr btn){ PostMessage(dlg, 0x0111, (IntPtr)GetDlgCtrlID(btn), btn); }   // WM_COMMAND/BN_CLICKED
  public static void CloseDlg(IntPtr dlg){ PostMessage(dlg, 0x0010, IntPtr.Zero, IntPtr.Zero); }                 // WM_CLOSE
  public static IntPtr FindClass(uint pid, string cls){
    IntPtr f=IntPtr.Zero;
    EnumWindows(delegate(IntPtr h, IntPtr l){
      uint p; GetWindowThreadProcessId(h, out p);
      if(p==pid && C(h)==cls){ f=h; return false; }
      return true; }, IntPtr.Zero);
    return f;
  }
  public static string Title(IntPtr h){ return T(h); }
  public static string BringOnScreen(IntPtr h){
    ShowWindow(h, 9);
    SetWindowPos(h, IntPtr.Zero, 80, 80, 1100, 720, 0x40);
    BringWindowToTop(h);
    SetForegroundWindow(h);
    return (GetForegroundWindow()==h) ? "fg" : "notfg";
  }
  public static bool IsForeground(IntPtr h){ return GetForegroundWindow()==h; }
}
'@

Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class WinAct {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  // 每行: hwnd|class|title|L,T,R,B  (仅可见的顶层窗口)
  public static List<string> VisibleTop(uint pid){
    var list=new List<string>();
    EnumWindows(delegate(IntPtr h, IntPtr l){
      uint p; GetWindowThreadProcessId(h, out p);
      if(p==pid && IsWindowVisible(h)){
        var c=new StringBuilder(256); GetClassName(h,c,256);
        var t=new StringBuilder(512); GetWindowText(h,t,512);
        RECT r; GetWindowRect(h, out r);
        list.Add(h.ToInt64() + "|" + c + "|" + t + "|" + r.L + "," + r.T + "," + r.R + "," + r.B);
      }
      return true; }, IntPtr.Zero);
    return list;
  }
  public static void ClickInside(IntPtr h, int dx, int dy){
    RECT r; GetWindowRect(h, out r);
    SetWindowPos(h, (IntPtr)(-1), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040);   // HWND_TOPMOST, NOMOVE|NOSIZE|SHOWWINDOW
    BringWindowToTop(h); SetForegroundWindow(h);
    int x = r.L + dx, y = r.T + dy;
    SetCursorPos(x, y);
    mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);      // LEFTDOWN
    mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);      // LEFTUP
  }
  // 当前前台窗口: pid / 类名 / 标题
  public static string Foreground(){
    IntPtr h = GetForegroundWindow(); uint p; GetWindowThreadProcessId(h, out p);
    var c = new StringBuilder(256); GetClassName(h, c, 256);
    var t = new StringBuilder(512); GetWindowText(h, t, 512);
    return "pid=" + p + " class=[" + c + "] title=[" + t + "]";
  }
}
'@

$script:policy = $Policy
$script:seen = @{}
$script:dlgLog = Join-Path $env:TEMP 'abp_ft_dialogs.log'
$script:resFile = Join-Path $env:TEMP 'abp_ft_results.txt'
if (Test-Path $script:dlgLog) { Remove-Item -Force $script:dlgLog }
if (Test-Path $script:resFile) { Remove-Item -Force $script:resFile }
$script:ppid = 0
$script:clicked = $false
$script:lock = Join-Path $env:TEMP 'abp_ft_pp.lock'     # 记录本编排器拉起的 PowerPoint PID
$script:policyFile = Join-Path $env:TEMP 'abp_ft_policy.txt'   # 测试驱动可逐次改写应答策略(yes/no/cancel/ok)
if (Test-Path $script:policyFile) { Remove-Item -Force $script:policyFile }

function Choose-Button($d) {
  # 返回所选按钮序号(与 d.Buttons 用 | 分隔的顺序一致), 找不到返回 -1
  $pol = $script:policy
  try { if (Test-Path $script:policyFile) { $t = ([IO.File]::ReadAllText($script:policyFile)).Trim(); if ($t) { $pol = $t.ToLower() } } } catch {}
  $names = $d.Buttons -split '\|'
  $end = -1; $ok = -1; $yes = -1; $no = -1; $cancel = -1; $dbg = -1; $cont = -1
  for ($i = 0; $i -lt $names.Length; $i++) {
    $n = $names[$i]
    if ($n -match '结束|&E\s*$|^End') { $end = $i }
    elseif ($n -match '^是|^Yes|&Y') { $yes = $i }
    elseif ($n -match '^否|^No\b|&N\s*$') { $no = $i }
    elseif ($n -match '取消|Cancel|&C\s*$') { $cancel = $i }
    elseif ($n -match '调试|&D\s*$|^Debug') { $dbg = $i }
    elseif ($n -match '继续|&C\s*$|^Continue') { $cont = $i }
    elseif ($n -match '确定|^OK') { $ok = $i }
  }
  if ($dbg -ge 0) {                                     # VBA 运行时错误框
    if ($ErrorBtn -eq 'debug' -and $dbg -ge 0) { return $dbg }
    if ($ErrorBtn -eq 'continue' -and $cont -ge 0) { return $cont }
    if ($end -ge 0) { return $end }
  }
  if ($pol -eq 'yes' -and $yes -ge 0) { return $yes }
  if ($pol -eq 'no' -and $no -ge 0) { return $no }
  if ($pol -eq 'cancel' -and $cancel -ge 0) { return $cancel }
  if ($ok -ge 0) { return $ok }
  if ($names.Length -gt 0) { return 0 }
  return -1
}

function Dismiss-Dlg($d) {
  $idx = Choose-Button $d
  $names = $d.Buttons -split '\|'
  if ($NoDismiss) { return '(record-only)' }
  if ($idx -ge 0 -and $idx -lt $d.Hs.Count) {
    [DlgSpy]::Click($d.H, $d.Hs[$idx])
    return $names[$idx]
  }
  [DlgSpy]::CloseDlg($d.H)
  return '(WM_CLOSE)'
}

function Add-DlgLine($line) {
  [IO.File]::AppendAllText($script:dlgLog, $line + "`r`n", (New-Object Text.UTF8Encoding($true)))
}

function Get-PaneInfo($cp, $tag) {
  $out = @()
  try {
    $sl = 0; $sc = 0; $el = 0; $ec = 0
    $cp.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
    $cm = $cp.CodeModule
    $proj = ''
    try { $proj = $cm.Parent.VBProject.FileName } catch {}
    $out += 'BREAK[' + $tag + '] module=[' + $cm.Parent.Name + '] project=[' + $proj + '] line=' + $sl
    $lo = [Math]::Max(1, $sl - 6); $hi = [Math]::Min($cm.CountOfLines, $sl + 2)
    for ($i = $lo; $i -le $hi; $i++) { $out += ('BREAK   ' + $(if ($i -eq $sl) { '>>' } else { '  ' }) + 'L' + $i + ': ' + $cm.Lines($i, 1)) }
  } catch { $out += 'BREAK[' + $tag + '] err: ' + $_.Exception.Message }
  return $out
}

function Read-BreakInfo {
  # 点了"调试"后, VBE 处于中断模式: 读出高亮的出错行(模块/行号/上下文); 活动窗格为空时扫描全部代码窗格
  # 每一步都留痕: 中断模式下外部 COM 能读到什么是不确定的, 空结果必须能区分"没读到"和"读错了"
  $lines = @()
  $a = $script:app
  $lines += 'BREAK step: app-null=' + ($a -eq $null)
  if ($a -eq $null) {
    try { $a = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application'); $lines += 'BREAK step: reattached' }
    catch { $lines += 'BREAK step: reattach err: ' + $_.Exception.Message }
  }
  $vbe = $null
  try { $vbe = $a.VBE; $lines += 'BREAK step: vbe-null=' + ($vbe -eq $null) } catch { $lines += 'BREAK step: VBE err: ' + $_.Exception.Message }
  if ($vbe -eq $null) { return $lines }
  $cp = $null
  for ($try = 0; $try -lt 6 -and $cp -eq $null; $try++) {
    try { $cp = $vbe.ActiveCodePane } catch { $lines += 'BREAK step: ActiveCodePane err: ' + $_.Exception.Message; break }
    if ($cp -eq $null) { Start-Sleep -Milliseconds 500 }
  }
  foreach ($d in @('Mode', 'ActiveVBProject.Name', 'SelectedVBComponent.Name', 'CodePanes.Count', 'ActiveWindow.Caption')) {
    try { $lines += 'BREAK diag ' + $d + '=' + (Invoke-Expression ('$vbe.' + $d)) } catch { $lines += 'BREAK diag ' + $d + ' err: ' + $_.Exception.Message }
  }
  $lines += 'BREAK step: activepane-null=' + ($cp -eq $null)
  if ($cp -ne $null) { $lines += (Get-PaneInfo $cp 'active'); return $lines }
  try {
    $n = $vbe.CodePanes.Count
    for ($k = 1; $k -le $n; $k++) { $lines += (Get-PaneInfo $vbe.CodePanes.Item($k) ('pane#' + $k)) }
  } catch { $lines += 'BREAK panes err: ' + $_.Exception.Message }
  return $lines
}

function Reset-Break {
  # 运行菜单 -> 重置 (Alt+R, 下, 下, 回车), 需要 VBE 窗口在前台
  $hv = [DlgSpy]::FindClass([uint32]$script:ppid, 'wndclass_desked_gsk')
  if ($hv -eq [IntPtr]::Zero) { return 'reset: no VBE window' }
  $fg = $false
  for ($i = 0; $i -lt 8 -and -not $fg; $i++) {
    [void][DlgSpy]::BringOnScreen($hv)
    Start-Sleep -Milliseconds 500
    $fg = [DlgSpy]::IsForeground($hv)
  }
  if (-not $fg) { return 'reset: VBE not foreground, keys NOT sent' }
  $wsh = New-Object -ComObject WScript.Shell
  $wsh.SendKeys('%r'); Start-Sleep -Milliseconds 600
  $wsh.SendKeys('{DOWN}{DOWN}{ENTER}'); Start-Sleep -Milliseconds 1500
  $t = [DlgSpy]::Title($hv)
  return ('reset: sent keys, VBE title=[' + $t + ']')
}

function Watch-Once {
  $found = @()
  foreach ($d in [DlgSpy]::Scan([uint32]$script:ppid)) {
    if (-not $script:seen.ContainsKey($d.H.ToInt64())) {
      $script:seen[$d.H.ToInt64()] = 1
      $line = 'DIALOG title=[' + $d.Title + '] text=[' + $d.Text + '] buttons=[' + $d.Buttons + ']'
      # 必须先落日志再点击: 点击后宏立即恢复执行, 驱动可能抢在日志落盘前读文件
      Add-DlgLine $line
      $found += $line
      $clicked = Dismiss-Dlg $d
      Add-DlgLine ('clicked=[' + $clicked + ']')
      if ($ShotOnDlg -and $d.Title -match 'Visual Basic' -and $d.Text -match '编译|语法|未定义') {
        # 点掉之后 VBE 会打开代码窗口并选中出错标识符 —— 此时截图
        Start-Sleep -Milliseconds 1800
        $shotScript = Join-Path $PSScriptRoot 'ppt-functest\dbg-shot.ps1'
        if (Test-Path $shotScript) {
          try {
            $so = & powershell -NoProfile -ExecutionPolicy Bypass -File $shotScript -Out ($ShotOnDlg + '.png') 2>&1
            Add-DlgLine ('SHOT: ' + (($so | Out-String).Trim() -replace "`r`n", ' | '))
          } catch { Add-DlgLine ('SHOT ERR: ' + $_.Exception.Message) }
        }
      }
      if ($clicked -match '调试') {
        Start-Sleep -Milliseconds 1500
        foreach ($b in (Read-BreakInfo)) { Add-DlgLine $b; $found += $b }
        if ($HoldOnBreak) {
          $go = Join-Path $env:TEMP 'abp_ft_go'
          if (Test-Path $go) { Remove-Item -Force $go }
          Add-DlgLine 'HOLD: break state kept; create %TEMP%\abp_ft_go to continue'
          $sw = [Diagnostics.Stopwatch]::StartNew()
          while (-not (Test-Path $go) -and $sw.Elapsed.TotalSeconds -lt 900) { Start-Sleep -Milliseconds 300 }
          Add-DlgLine 'HOLD: released'
        }
        $r = Reset-Break
        Add-DlgLine $r; $found += $r
        Start-Sleep -Milliseconds 800
      }
    }
  }
  if ($ClickTitle -and -not $script:clicked) {
    $all = [WinAct]::VisibleTop([uint32]$script:ppid)
    foreach ($w in $all) {
      $parts = $w -split '\|'
      if ($parts[2] -match $ClickTitle) {
        $script:clicked = $true
        foreach ($x in $all) { $line = 'WINDOW ' + $x; Add-DlgLine $line; $found += $line }
        [WinAct]::ClickInside([IntPtr][int64]$parts[0], 40, 45)
        $line = 'CLICK window=[' + $parts[2] + '] class=[' + $parts[1] + '] rect=[' + $parts[3] + ']'
        Add-DlgLine $line; $found += $line
        Start-Sleep -Milliseconds 500
        $line = 'FOREGROUND ' + [WinAct]::Foreground()
        Add-DlgLine $line; $found += $line
        break
      }
    }
  }
  return $found
}

function Invoke-Macro([string]$entry, [string[]]$margs = @(), [int]$timeoutSec = 120) {
  $job = Start-Job -Name ('abp' + [Guid]::NewGuid().ToString('N')) -ScriptBlock {
    param($entry, $margs)
    $app = $null
    for ($i = 0; $i -lt 20 -and $app -eq $null; $i++) {
      try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { Start-Sleep -Milliseconds 250 }
    }
    if ($app -eq $null) { 'ERR:no PowerPoint in ROT'; return }
    $a = New-Object System.Collections.ArrayList
    [void]$a.Add($entry); foreach ($x in $margs) { [void]$a.Add($x) }
    try { $r = $app.GetType().InvokeMember('Run', 'InvokeMethod', $null, $app, $a.ToArray()); 'RET:[' + $r + ']' }
    catch { 'ERR:' + $_.Exception.Message }
  } -ArgumentList $entry, $margs
  $sw = [Diagnostics.Stopwatch]::StartNew(); $timedOut = $false
  while ($job.State -eq 'Running') {
    [void](Watch-Once)
    if ($sw.Elapsed.TotalSeconds -gt $timeoutSec) { $timedOut = $true; break }
    Start-Sleep -Milliseconds 120
  }
  $out = ''
  if (-not $timedOut) { $out = (Receive-Job $job | Out-String).Trim() }
  Remove-Job $job -Force -ErrorAction SilentlyContinue
  if ($timedOut) { return 'TIMEOUT after ' + $timeoutSec + 's' }
  return $out
}

function Wait-Exit([int]$sec) {
  $sw = [Diagnostics.Stopwatch]::StartNew()
  while ($sw.Elapsed.TotalSeconds -lt $sec -and (Get-Process POWERPNT -ErrorAction SilentlyContinue)) { Start-Sleep -Milliseconds 300 }
  return -not [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
}

function Recover-And-Quit($app) {
  # PowerPoint 拒绝 Quit 通常是 VBA 停在中断模式: 还原 VBE 窗口 -> 运行菜单 -> 重置 -> 再 Quit
  $note = @()
  if (-not (Get-Process POWERPNT -ErrorAction SilentlyContinue)) { return $note }
  $hv = [DlgSpy]::FindClass([uint32]$script:ppid, 'wndclass_desked_gsk')
  $vt = ''
  if ($hv -ne [IntPtr]::Zero) { $vt = [DlgSpy]::Title($hv) }
  $note += 'recover: VBE hwnd=' + $hv + ' title=[' + $vt + ']'
  if ($vt -notmatch '\[.+\]' -or $vt -notmatch '\[(\u4e2d\u65ad|break)\]') {
    $note += 'recover: VBE not in break mode -> no keys sent'      # 非中断状态绝不能发菜单快捷键
  }
  elseif ($hv -ne [IntPtr]::Zero) {
    $fg = $false
    for ($i = 0; $i -lt 8 -and -not $fg; $i++) {
      $note += 'recover: bring=' + [DlgSpy]::BringOnScreen($hv)
      Start-Sleep -Milliseconds 700
      $fg = [DlgSpy]::IsForeground($hv)
    }
    if ($fg) {
      $wsh = New-Object -ComObject WScript.Shell
      $wsh.SendKeys('%r'); Start-Sleep -Milliseconds 600
      $wsh.SendKeys('{DOWN}{DOWN}{ENTER}'); Start-Sleep -Milliseconds 1500
      $note += 'recover: sent Alt+R,Down,Down,Enter'
    } else { $note += 'recover: VBE not foreground, keys not sent' }
  }
  try {
    $n = $app.Presentations.Count
    for ($i = 0; $i -lt $n; $i++) { $p = $app.Presentations.Item(1); $p.Saved = -1; $p.Close() }
    $app.Quit(); $note += 'recover: Quit OK'
  } catch { $note += ('recover: COM ERR ' + $_.Exception.Message) }
  $note += 'recover: exited=' + (Wait-Exit 25)
  return $note
}

# ---------------- main ----------------
# 严格串行: 上一轮由本编排器拉起的 PowerPoint 若还在退出中, 先等它彻底消失, 绝不挂到将死的实例上
if (Test-Path $script:lock) {
  $lockPid = 0
  try { $lockPid = [int]((Get-Content $script:lock -Raw).Trim()) } catch {}
  if ($lockPid -gt 0 -and (Get-Process -Id $lockPid -ErrorAction SilentlyContinue)) {
    Write-Host ('waiting for previous test PowerPoint (pid ' + $lockPid + ') to exit...')
    [void](Wait-Exit 90)
  }
  Remove-Item -Force $script:lock -ErrorAction SilentlyContinue
}
$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$app = New-Object -ComObject PowerPoint.Application
$script:app = $app
$app.AutomationSecurity = 1
try { $app.Visible = -1 } catch {}
Start-Sleep -Milliseconds 800
# COM 拉起的 PowerPoint 主窗口默认是最小化的(屏幕外), 弹窗/面板的 UIA 与截图都依赖它在屏内
try { $app.WindowState = 3 } catch {}          # ppWindowMaximize
$script:ppid = (Get-Process POWERPNT | Select-Object -First 1).Id
if (-not $wasRunning) { [IO.File]::WriteAllText($script:lock, [string]$script:ppid) }

$utf8 = New-Object Text.UTF8Encoding($true)
$code = (([IO.File]::ReadAllText((Resolve-Path -LiteralPath $Driver).Path, $utf8)) -split '\r?\n' | Where-Object { $_ -notmatch '^Attribute\s' }) -join "`r`n"

$pres = $null
foreach ($f in $OpenFirst) {
  $null = $app.Presentations.Open((Resolve-Path -LiteralPath $f).Path, 0, 0, -1)
  Start-Sleep -Milliseconds 500
}
if ($NoScratch) { $pres = $app.ActivePresentation } else { $pres = $app.Presentations.Add(-1) }
$mod = $pres.VBProject.VBComponents.Add(1)
$mod.Name = 'modFTDriver'
$mod.CodeModule.AddFromString($code)
# 回读 VBA 实际收到的模块文本(排查 AddFromString 编码/截断问题), 存 out\<tag>.mod-<名>.txt
$dumpDir = Join-Path $PSScriptRoot 'ppt-functest\out'
New-Item -ItemType Directory -Force -Path $dumpDir | Out-Null
try { [IO.File]::WriteAllText((Join-Path $dumpDir ($Tag + '.mod-' + $mod.Name + '.txt')), $mod.CodeModule.Lines(1, $mod.CodeModule.CountOfLines), $utf8) } catch {}
if ($SimpleBas) {
  $sc = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $SimpleBas).Path, $utf8)
  $mname = 'modSimple'
  if ($sc -match 'Attribute\s+VB_Name\s*=\s*"([^"]+)"') { $mname = $Matches[1] }
  $sc = (($sc -split '\r?\n') | Where-Object { $_ -notmatch '^Attribute\s' }) -join "`r`n"
  $mod2 = $pres.VBProject.VBComponents.Add(1)
  $mod2.Name = $mname
  $mod2.CodeModule.AddFromString($sc)
  try { [IO.File]::WriteAllText((Join-Path $dumpDir ($Tag + '.mod-' + $mname + '.txt')), $mod2.CodeModule.Lines(1, $mod2.CodeModule.CountOfLines), $utf8) } catch {}
}
$notes = @('DRIVER injected -> ' + $pres.Name + '  entry=' + $Entry + '  policy=' + $Policy + '  ppid=' + $script:ppid)

$ret = Invoke-Macro $Entry @() $TimeoutSec
$notes += '=== RUN ' + $Entry + ' => ' + $ret
if ($ret -like 'TIMEOUT*') {
  Start-Sleep -Milliseconds 500
  [void](Watch-Once)     # 超时多为被弹窗挡住: 再清一次, 给 PowerPoint 自己收尾的机会
}
Start-Sleep -Milliseconds 400
[void](Watch-Once)

# 汇总输出: 控制台是 GBK, 中文会乱码, 所以落一份 UTF-8 日志供查阅
$outLines = New-Object System.Collections.ArrayList
foreach ($n in $notes) { [void]$outLines.Add($n) }
[void]$outLines.Add('--- dialogs ---')
if (Test-Path $script:dlgLog) { foreach ($l in (Get-Content $script:dlgLog -Encoding UTF8)) { [void]$outLines.Add($l) } } else { [void]$outLines.Add('(none)') }
[void]$outLines.Add('--- results ---')
if (Test-Path $script:resFile) { foreach ($l in (Get-Content $script:resFile -Encoding Default)) { [void]$outLines.Add($l) } } else { [void]$outLines.Add('(no result file)') }

# cleanup
try {
  if ($NoScratch) { $pres.VBProject.VBComponents.Remove($mod) }
  $pres.Saved = -1; $pres.Close()
} catch { $outLines.Add('cleanup: ' + $_.Exception.Message) }
if (-not $wasRunning -and -not $KeepOpen) {
  # 带无模式窗体/定时器时 PowerPoint 退出极慢: 先让加载项自己收尾
  $ac = Invoke-Macro 'Auto_Close' @() 25
  [void]$outLines.Add('cleanup: Auto_Close => ' + $ac)
  try { $app.Quit() } catch {}
  # 编排器自己持有的 COM 引用会让 POWERPNT 进程无法结束: 必须先全部释放
  $pres = $null; $mod = $null; $app = $null; $script:app = $null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
  if (-not (Wait-Exit 45)) {
    $app2 = $null
    try { $app2 = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch {}
    if ($app2 -ne $null) {
      foreach ($n in (Recover-And-Quit $app2)) { [void]$outLines.Add($n) }
      try { $app2.Quit() } catch {}
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app2)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [void]$outLines.Add('cleanup: POWERPNT alive=' + (-not (Wait-Exit 30)))
  }
  if (-not (Get-Process POWERPNT -ErrorAction SilentlyContinue)) { Remove-Item -Force $script:lock -ErrorAction SilentlyContinue }
}
$outDir = Join-Path $PSScriptRoot 'ppt-functest\out'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$outLog = Join-Path $outDir ($Tag + '.log')
[IO.File]::WriteAllLines($outLog, $outLines, (New-Object Text.UTF8Encoding($true)))
$pass = @($outLines | Where-Object { $_ -match '^PASS' }).Count
$fail = @($outLines | Where-Object { $_ -match '^FAIL' }).Count
Write-Host ('PASS=' + $pass + ' FAIL=' + $fail + '  dialogs=' + @($outLines | Where-Object { $_ -match '^DIALOG' }).Count + '  log=' + $Tag + '.log')
Write-Host 'FT-END'
