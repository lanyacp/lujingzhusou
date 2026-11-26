Attribute VB_Name = "modFTDriver"
Option Explicit

' =====================================================================
'  路径助手 PowerPoint 版 功能测试驱动
'  由 tools/ppt-functest.ps1 注入临时演示文稿, 再调用入口宏:
'    FT_All          主套件(容错/直线/锚点/中心十字/航点/抛物线/圆弧/描迹/过期图形/简易宏)
'    FT_PanelStress  面板反复显示/隐藏 + 高频读面板状态(针对"运行时错误9"的回归)
'    FT_PreStart     会直接打死驱动的用例(未点开始就点"清空航点"), 单独跑
'  结果: 每行 PASS / FAIL / INFO, 写 %TEMP%\abp_ft_results.txt
'        弹窗由编排器的看门狗记录到 abp_ft_dialogs.log 并自动关闭
'  加载项里的动作通过 Application.Run 调用(与真人按 Alt+F8 / 功能区回调同一条路)
' =====================================================================

Private Const MK As String = "__pthlp_"
Private Const NM_A As String = "__pthlp_A__"
Private Const NM_B As String = "__pthlp_B__"
Private Const NM_PC As String = "__pthlp_PC__"
Private Const NM_PD As String = "__pthlp_PD__"
Private Const NM_PM As String = "__pthlp_PM__"
Private Const NM_TR As String = "__pthlp_TR__"

' 移动图形 (100,100,120,80) 中心 O=(160,140); 目标 (500,300,100,100) 中心 T=(550,350)
Private Const OX As Single = 160
Private Const OY As Single = 140
Private Const TX As Single = 550
Private Const TY As Single = 350

Private gPass As Long, gFail As Long
Private gS As Slide
Private gMover As Shape, gTarget As Shape
Private gSW As Single, gSH As Single

' ---------------- 基础工具 ----------------

Private Sub W(ByVal s As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open Environ$("TEMP") & "\abp_ft_results.txt" For Append As #f
    Print #f, s
    Close #f
End Sub

Private Sub Chk(ByVal nm As String, ByVal cond As Boolean, ByVal detail As String)
    If cond Then gPass = gPass + 1: W "PASS  " & nm & "  " & detail Else gFail = gFail + 1: W "FAIL  " & nm & "  " & detail
End Sub

Private Sub Info(ByVal s As String)
    W "INFO  " & s
End Sub

Private Function Fmt(ByVal v As Single) As String
    Fmt = Format$(v, "0.000000")
End Function

Private Sub Pause(ByVal sec As Single)
    Dim t As Single: t = Timer
    Do While Timer < t + sec
        DoEvents
    Loop
End Sub

Private Function CountStr(ByVal hay As String, ByVal needle As String) As Long
    Dim p As Long, n As Long
    p = InStr(1, hay, needle)
    Do While p > 0
        n = n + 1
        p = InStr(p + Len(needle), hay, needle)
    Loop
    CountStr = n
End Function

' 看门狗弹窗日志全文(UTF-8, 这里只用 ASCII 片段判断)
Private Function DlgText() As String
    Dim f As Integer, i As Long, fn As String
    fn = Environ$("TEMP") & "\abp_ft_dialogs.log"
    For i = 1 To 10
        Err.Clear
        On Error Resume Next
        f = FreeFile
        Open fn For Input As #f
        If Err.Number = 0 Then
            DlgText = Input(LOF(f), f)
            Close #f
            Exit Function
        End If
        Pause 0.1
    Next i
End Function

' 调用加载项宏; 返回 ok / ERRn / VBAERRDIALOG(弹了 VBA 原始错误框) / DIALOG+n(弹了 n 个提示框)
Private Sub EnsureDlgFile()
    On Error Resume Next
    Dim t As Integer: t = FreeFile
    Open Environ$("TEMP") & "bp_ft_dialogs.log" For Append As #t
    Close #t
End Sub

Private Function RunM(ByVal m As String) As String
    Dim d0 As Long, r0 As Long, e As Long, ds As String, dlgA As String, dN As Long, rN As Long
    EnsureDlgFile
    dlgA = DlgText()
    d0 = CountStr(dlgA, "DIALOG title=[")
    r0 = CountStr(dlgA, "title=[Microsoft Visual Basic]")
    Err.Clear
    On Error Resume Next
    Application.Run m
    e = Err.Number: ds = Err.Description
    On Error GoTo 0
    Pause 0.12
    dlgA = DlgText()
    dN = CountStr(dlgA, "DIALOG title=[") - d0
    rN = CountStr(dlgA, "title=[Microsoft Visual Basic]") - r0
    If e <> 0 Then
        RunM = "ERR" & e & ":" & ds
    ElseIf rN > 0 Then
        RunM = "VBAERRDIALOG"
    ElseIf dN > 0 Then
        RunM = "DIALOG+" & dN
    Else
        RunM = "ok"
    End If
End Function

Private Sub ExpectOk(ByVal nm As String, ByVal m As String)
    Dim r As String: r = RunM(m)
    Chk nm, (r = "ok"), m & " => " & r
End Sub

' 逐次设置看门狗对 是/否/取消 的应答
Private Sub SetPolicy(ByVal p As String)
    Dim f As Integer
    f = FreeFile
    Open Environ$("TEMP") & "\abp_ft_policy.txt" For Output As #f
    Print #f, p
    Close #f
End Sub

' ---------------- 图形 / 标记 ----------------

Private Function FindS(ByVal nm As String) As Shape
    Dim i As Long
    For i = 1 To gS.Shapes.Count
        If gS.Shapes(i).Name = nm Then Set FindS = gS.Shapes(i): Exit Function
    Next i
End Function

Private Function CountS(ByVal nm As String) As Long
    Dim i As Long
    For i = 1 To gS.Shapes.Count
        If gS.Shapes(i).Name = nm Then CountS = CountS + 1
    Next i
End Function

Private Function Cx(ByVal sh As Shape) As Single
    Cx = sh.Left + sh.Width / 2!
End Function

Private Function Cy(ByVal sh As Shape) As Single
    Cy = sh.Top + sh.Height / 2!
End Function

Private Function MarkAt(ByVal nm As String, ByVal x As Single, ByVal y As Single, ByVal tol As Single) As Boolean
    Dim sh As Shape: Set sh = FindS(nm)
    If sh Is Nothing Then Exit Function
    MarkAt = (Abs(Cx(sh) - x) <= tol And Abs(Cy(sh) - y) <= tol)
End Function

Private Sub MoveMark(ByVal nm As String, ByVal x As Single, ByVal y As Single)
    Dim sh As Shape: Set sh = FindS(nm)
    If Not sh Is Nothing Then sh.Left = x - sh.Width / 2!: sh.Top = y - sh.Height / 2!
End Sub

Private Function MarkPos(ByVal nm As String) As String
    Dim sh As Shape: Set sh = FindS(nm)
    If sh Is Nothing Then MarkPos = "(none)" Else MarkPos = "(" & Fmt(Cx(sh)) & "," & Fmt(Cy(sh)) & ")"
End Function

Private Sub SelMover()
    gMover.Select
End Sub

Private Sub SelTarget()
    gTarget.Select
End Sub

Private Sub SelNone()
    On Error Resume Next
    ActiveWindow.Selection.Unselect                ' 真正取消图形选择
End Sub

' ---------------- 动画回读 ----------------

Private Function NEff() As Long
    NEff = gS.TimeLine.MainSequence.Count
End Function

Private Function PathPts(ByVal pth As String, ByRef n As Long) As Variant
    Dim t() As String, i As Long, k As Long, xs() As Single
    n = 0
    If Len(pth) = 0 Then Exit Function
    t = Split(Replace(Replace(Replace(pth, ",", " "), vbCr, " "), vbLf, " "), " ")
    ReDim xs(0 To 400)
    k = 0
    For i = 0 To UBound(t)
        If t(i) = "M" Or t(i) = "L" Then
            If k + 2 > 400 Then Exit For
            xs(k) = CSng(t(i + 1)): xs(k + 1) = CSng(t(i + 2)): k = k + 2: n = n + 1
        End If
    Next i
    PathPts = xs
End Function

' 读取第 idx 个动画的运动路径: 返回点数; pts(i)=Array(dx,dy) 为相对起点的位移(磅)
Private Function ReadPath(ByVal idx As Long, ByRef pts() As Variant, ByRef shp As String, ByRef dur As Single, ByRef trig As Long) As Long
    Dim eff As Effect, i As Long, beh As AnimationBehavior, mb As AnimationBehavior
    Dim pth As String, raw As Variant, n As Long, k As Long
    Set eff = gS.TimeLine.MainSequence(idx)
    shp = eff.Shape.Name: dur = eff.Timing.Duration: trig = eff.Timing.TriggerType
    For i = 1 To eff.Behaviors.Count
        Set beh = eff.Behaviors(i)
        If beh.Type = msoAnimTypeMotion Then Set mb = beh
    Next i
    If mb Is Nothing Then Exit Function
    pth = mb.MotionEffect.Path
    If Len(pth) = 0 Then         ' 用 From/To 建的直线
        ReDim pts(0 To 1)
        pts(0) = Array(mb.MotionEffect.FromX * gSW, mb.MotionEffect.FromY * gSH)
        pts(1) = Array(mb.MotionEffect.ToX * gSW, mb.MotionEffect.ToY * gSH)
        ReadPath = 2
        Exit Function
    End If
    raw = PathPts(pth, n)
    If n = 0 Then Exit Function
    ReDim pts(0 To n - 1)
    For k = 0 To n - 1
        pts(k) = Array(CSng(raw(2 * k)) * gSW, CSng(raw(2 * k + 1)) * gSH)
    Next k
    ReadPath = n
End Function

' 路径末点与期望位移(dx,dy 磅)的误差
Private Function EndErr(ByRef pts() As Variant, ByVal n As Long, ByVal dx As Single, ByVal dy As Single) As Single
    EndErr = Sqr((pts(n - 1)(0) - dx) ^ 2 + (pts(n - 1)(1) - dy) ^ 2)
End Function

' ---------------- 场景 ----------------


Private Sub InitScene()
    Dim pres As Presentation: Set pres = ActivePresentation
    Do While pres.Slides.Count > 1
        pres.Slides(pres.Slides.Count).Delete
    Loop
    If pres.Slides.Count = 0 Then
        pres.Slides.Add 1, ppLayoutBlank
    End If
    Set gS = pres.Slides(1)
    Do While gS.Shapes.Count > 0
        gS.Shapes(1).Delete
    Loop
    Do While NEff() > 0
        gS.TimeLine.MainSequence(1).Delete
    Loop
    gSW = pres.PageSetup.SlideWidth: gSH = pres.PageSetup.SlideHeight
    Set gMover = gS.Shapes.AddShape(msoShapeRectangle, 100!, 100!, 120!, 80!)
    gMover.Name = "FT_mover"
    Set gTarget = gS.Shapes.AddShape(msoShapeOval, 500!, 300!, 100!, 100!)
    gTarget.Name = "FT_target"
End Sub

' ---------------- 套件 ----------------

Public Sub FT_All()
    gPass = 0: gFail = 0
    On Error GoTo eh
    On Error Resume Next
    Kill Environ$("TEMP") & "\abp_ft_results.txt"
    On Error GoTo 0
    Dim ex As Single, ey As Single
    InitScene
    Info "slide " & gSW & "x" & gSH & "  O=(" & OX & "," & OY & ") T=(" & TX & "," & TY & ")"
    ex = (TX - OX) / gSW: ey = (TY - OY) / gSH

    ' ===== 0. 未点"开始"就点各按钮: 只应更新状态栏, 绝不能弹错误框 =====
    Dim n0 As Long: n0 = NEff()
    ExpectOk "pre-start 放C点", "actPlaceC"
    ExpectOk "pre-start 放D点", "actPlaceD"
    ExpectOk "pre-start 放P点", "actPlaceP"
    ExpectOk "pre-start +航点", "actAddW"
    ExpectOk "pre-start 开始描迹", "actTraceStart"
    ExpectOk "pre-start 记一点", "actTracePoint"
    ExpectOk "pre-start 完成描迹", "actTraceDone"
    ExpectOk "pre-start 生成", "actGenerate"
    ExpectOk "pre-start 清理标记", "actClean"
    SelNone
    ExpectOk "pre-start 开始(无选择)", "actStart"
    ExpectOk "pre-start 中心十字(无选择)", "actCenterCross"
    Chk "pre-start 不产生动画", NEff() = n0, "effects " & n0 & " -> " & NEff()
    Chk "pre-start 不产生标记", gS.Shapes.Count = 2, "shapes=" & gS.Shapes.Count

    ' ===== 1. 直线: 中心到中心 =====
    SelMover
    ExpectOk "line 开始", "actStart"
    Chk "line 放置 A/B 标记", (CountS(NM_A) = 1 And CountS(NM_B) = 1), "A=" & CountS(NM_A) & " B=" & CountS(NM_B)
    Chk "line A 在图形中心", MarkAt(NM_A, OX, OY, 0.01), "A=" & MarkPos(NM_A)
    Chk "line B 初始在图形中心", MarkAt(NM_B, OX, OY, 0.01), ""
    n0 = NEff()
    MoveMark NM_B, TX, TY
    ExpectOk "line 生成", "actGenerate"
    Dim pts() As Variant, shp As String, dur As Single, trig As Long, n As Long
    Chk "line 新增 1 个动画", NEff() = n0 + 1, "effects " & n0 & " -> " & NEff()
    If NEff() > n0 Then
        n = ReadPath(NEff(), pts, shp, dur, trig)
        Chk "line 动画挂在移动图形上", shp = "FT_mover", "shape=" & shp
        Chk "line 起点为原位", (n >= 2 And Abs(pts(0)(0)) < 0.01 And Abs(pts(0)(1)) < 0.01), "start=(" & Fmt(pts(0)(0)) & "," & Fmt(pts(0)(1)) & ")"
        Chk "line 终点=目标中心(精确)", EndErr(pts, n, TX - OX, TY - OY) < 0.01, "end err=" & Fmt(EndErr(pts, n, TX - OX, TY - OY)) & "pt  expect=(" & Fmt(TX - OX) & "," & Fmt(TY - OY) & ")"
        Chk "line 时长=2s", Abs(dur - 2!) < 0.01, "dur=" & dur
        Chk "line 触发=单击时", trig = msoAnimTriggerOnPageClick, "trig=" & trig
    End If
    Chk "line 生成后标记已清理", (CountS(NM_A) = 0 And CountS(NM_B) = 0), "A=" & CountS(NM_A) & " B=" & CountS(NM_B)

    ' ===== 2. 吸附: B 放在目标中心附近(偏差 8pt/-6pt) =====
    SelMover
    ExpectOk "snap 开始", "actStart"
    MoveMark NM_B, TX + 8!, TY - 6!
    n0 = NEff()
    ExpectOk "snap 生成", "actGenerate"
    If NEff() > n0 Then
        n = ReadPath(NEff(), pts, shp, dur, trig)
        Chk "snap 终点吸附到精确中心", EndErr(pts, n, TX - OX, TY - OY) < 0.01, "end err=" & Fmt(EndErr(pts, n, TX - OX, TY - OY)) & "pt"
    Else
        Chk "snap 新增动画", False, "effects " & n0 & " -> " & NEff()
    End If

    ' ===== 3. 起点终点重合 =====
    SelMover
    ExpectOk "zero 开始", "actStart"
    n0 = NEff()
    Dim rz As String: rz = RunM("actGenerate")
    Chk "zero 零位移: 仅提示不产生动画", (rz = "ok" And NEff() = n0), rz & "  effects " & n0 & "->" & NEff()

    ' ===== 4. 自定义锚点 C -> D =====
    SelMover
    ExpectOk "anchor 开始", "actStart"
    SelMover
    ExpectOk "anchor 放C点", "actPlaceC"
    Chk "anchor C 标记存在且唯一", CountS(NM_PC) = 1, "count=" & CountS(NM_PC)
    MoveMark NM_PC, OX - 50!, OY - 30!           ' 锚点: 图形中心左上 (110,110)
    SelTarget
    ExpectOk "anchor 放D点(选中目标)", "actPlaceD"
    Chk "anchor D 落在所选目标中心", MarkAt(NM_PD, TX, TY, 0.01), "D=" & MarkPos(NM_PD)
    n0 = NEff()
    ExpectOk "anchor 生成", "actGenerate"
    If NEff() > n0 Then
        n = ReadPath(NEff(), pts, shp, dur, trig)
        Chk "anchor 位移 = D - C (精确)", EndErr(pts, n, TX - (OX - 50!), TY - (OY - 30!)) < 0.01, "end err=" & Fmt(EndErr(pts, n, TX - (OX - 50!), TY - (OY - 30!))) & "pt"
    Else
        Chk "anchor 新增动画", False, "effects " & n0 & " -> " & NEff()
    End If

    ' ===== 5. 中心十字 + 吸附到十字 =====
    SelTarget
    ExpectOk "cross 加中心十字", "actCenterCross"
    Dim cnm As String: cnm = MK & "C_" & gTarget.Id
    Chk "cross 十字在目标中心", MarkAt(cnm, TX, TY, 0.01), "name=" & cnm
    SelMover
    ExpectOk "cross-snap 开始", "actStart"
    MoveMark NM_B, TX + 10!, TY
    n0 = NEff()
    ExpectOk "cross-snap 生成", "actGenerate"
    If NEff() > n0 Then
        n = ReadPath(NEff(), pts, shp, dur, trig)
        Chk "cross-snap 终点吸附到十字", EndErr(pts, n, TX - OX, TY - OY) < 0.01, "end err=" & Fmt(EndErr(pts, n, TX - OX, TY - OY)) & "pt"
    Else
        Chk "cross-snap 新增动画", False, "effects " & n0 & " -> " & NEff()
    End If
    Chk "cross-snap 十字未被清理(吸附目标保留)", CountS(cnm) = 1, "count=" & CountS(cnm)
    SelTarget
    ExpectOk "cross 移除中心十字", "actCenterCross"
    Chk "cross 十字已移除", CountS(cnm) = 0, "count=" & CountS(cnm)

    ' ===== 6. 航点 =====
    SelMover
    Application.Run "SetForm", "航点"
    ExpectOk "wp 开始", "actStart"
    Chk "wp 航点形态不生成 B 标记", CountS(NM_B) = 0, "B=" & CountS(NM_B)
    SelTarget
    ExpectOk "wp +航点1(选中目标)", "actAddW"
    Chk "wp W1 在目标中心", MarkAt(MK & "W1__", TX, TY, 0.01), "count W1=" & CountS(MK & "W1__")
    SelNone
    ExpectOk "wp +航点2(默认偏移)", "actAddW"
    Chk "wp W2 编号递增且在默认位置", (CountS(MK & "W2__") = 1 And MarkAt(MK & "W2__", OX + 52!, OY + 28!, 0.01)), "W2 count=" & CountS(MK & "W2__")
    SelTarget
    ExpectOk "wp +航点3(选中目标)", "actAddW"
    Chk "wp W3 编号递增", CountS(MK & "W3__") = 1, "W3 count=" & CountS(MK & "W3__")
    n0 = NEff()
    ExpectOk "wp 生成", "actGenerate"
    Chk "wp 生成新增 1 个动画(一条折线)", NEff() = n0 + 1, "effects " & n0 & " -> " & NEff()
    If NEff() > n0 Then
        n = ReadPath(NEff(), pts, shp, dur, trig)
        Chk "wp 动画挂在移动图形上", shp = "FT_mover", "shape=" & shp
        Chk "wp 路径点数=4 (起点,W1,W2,W3)", n = 4, "n=" & n
        If n = 4 Then
            Chk "wp W1", (Abs(pts(1)(0) - (TX - OX)) < 0.01 And Abs(pts(1)(1) - (TY - OY)) < 0.01), "(" & Fmt(pts(1)(0)) & "," & Fmt(pts(1)(1)) & ")"
            Chk "wp W2", (Abs(pts(2)(0) - 52!) < 0.01 And Abs(pts(2)(1) - 28!) < 0.01), "(" & Fmt(pts(2)(0)) & "," & Fmt(pts(2)(1)) & ")"
            Chk "wp W3=终点", (Abs(pts(3)(0) - (TX - OX)) < 0.01 And Abs(pts(3)(1) - (TY - OY)) < 0.01), "(" & Fmt(pts(3)(0)) & "," & Fmt(pts(3)(1)) & ")"
        End If
        Chk "wp 整段时长=2s", Abs(dur - 2!) < 0.01, "dur=" & dur
    End If

    ' ===== 7. 抛物线 =====
    SelMover
    Application.Run "SetForm", "抛物线"
    ExpectOk "para 开始", "actStart"
    SelNone
    ExpectOk "para 放P(顶点)", "actPlaceP"
    Chk "para P 标记存在", CountS(NM_PM) = 1, "count=" & CountS(NM_PM)
    MoveMark NM_PM, 300!, 60!                    ' 顶点
    SelTarget
    ExpectOk "para 放D(终点)", "actPlaceD"
    n0 = NEff()
    ExpectOk "para 生成", "actGenerate"
    Chk "para 生成新增 1 个动画", NEff() = n0 + 1, "effects " & n0 & " -> " & NEff()
    If NEff() > n0 Then
        n = ReadPath(NEff(), pts, shp, dur, trig)
        Chk "para 起点为原位", (Abs(pts(0)(0)) < 0.01 And Abs(pts(0)(1)) < 0.01), ""
        Chk "para 终点=目标中心(精确)", EndErr(pts, n, TX - OX, TY - OY) < 0.01, "end err=" & Fmt(EndErr(pts, n, TX - OX, TY - OY)) & "pt  n=" & n
        Dim k As Long, best As Single: best = 1E+9
        For k = 0 To n - 1
            If Sqr((pts(k)(0) - 140!) ^ 2 + (pts(k)(1) + 80!) ^ 2) < best Then best = Sqr((pts(k)(0) - 140!) ^ 2 + (pts(k)(1) + 80!) ^ 2)
        Next k
        Chk "para 精确经过顶点", best < 0.05, "vertex err=" & Fmt(best) & "pt"
        Chk "para 采样 13 点(12 段)", n = 13, "n=" & n
    End If

    ' ===== 8. 圆弧 =====
    SelMover
    Application.Run "SetForm", "圆弧"
    ExpectOk "arc 开始", "actStart"
    SelNone
    ExpectOk "arc 放M(弧点)", "actPlaceP"
    MoveMark NM_PM, 420!, 120!                   ' 弧上一点
    SelTarget
    ExpectOk "arc 放D(终点)", "actPlaceD"
    n0 = NEff()
    ExpectOk "arc 生成", "actGenerate"
    Chk "arc 生成新增 1 个动画", NEff() = n0 + 1, "effects " & n0 & " -> " & NEff()
    If NEff() > n0 Then
        n = ReadPath(NEff(), pts, shp, dur, trig)
        Chk "arc 终点=目标中心(精确)", EndErr(pts, n, TX - OX, TY - OY) < 0.01, "end err=" & Fmt(EndErr(pts, n, TX - OX, TY - OY)) & "pt  n=" & n
        best = 1E+9
        For k = 0 To n - 1
            If Sqr((pts(k)(0) - 260!) ^ 2 + (pts(k)(1) + 20!) ^ 2) < best Then best = Sqr((pts(k)(0) - 260!) ^ 2 + (pts(k)(1) + 20!) ^ 2)
        Next k
        Chk "arc 经过弧点(采样间距内)", best < 8!, "mid err=" & Fmt(best) & "pt (25采样, R=244)"
        ' 共圆: 用 起点/中点/终点 求圆心, 所有点到圆心距离应一致
        Dim ax As Double, ay As Double, bx As Double, byy As Double, cx2 As Double, cy2 As Double
        Dim d As Double, ux As Double, uy As Double, r0 As Double, dev As Single, rr As Double
        ax = pts(0)(0): ay = pts(0)(1)
        bx = pts(n \ 2)(0): byy = pts(n \ 2)(1)
        cx2 = pts(n - 1)(0): cy2 = pts(n - 1)(1)
        d = 2 * (ax * (byy - cy2) + bx * (cy2 - ay) + cx2 * (ay - byy))
        If Abs(d) < 0.000001 Then
            Chk "arc 三点不共线(应是圆弧)", False, "collinear"
        Else
            ux = ((ax * ax + ay * ay) * (byy - cy2) + (bx * bx + byy * byy) * (cy2 - ay) + (cx2 * cx2 + cy2 * cy2) * (ay - byy)) / d
            uy = ((ax * ax + ay * ay) * (cx2 - bx) + (bx * bx + byy * byy) * (ax - cx2) + (cx2 * cx2 + cy2 * cy2) * (bx - ax)) / d
            r0 = Sqr((ax - ux) ^ 2 + (ay - uy) ^ 2)
            dev = 0!
            For k = 0 To n - 1
                rr = Abs(Sqr((pts(k)(0) - ux) ^ 2 + (pts(k)(1) - uy) ^ 2) - r0)
                If rr > dev Then dev = CSng(rr)
            Next k
            Chk "arc 所有点共圆", dev < 0.05, "max dev=" & Fmt(dev) & "pt  R=" & Fmt(CSng(r0))
        End If
    End If

    ' ===== 9. 描迹(含定时器自动记点) =====
    SelMover
    Application.Run "SetForm", "描迹"
    ExpectOk "trace 开始", "actStart"
    SelNone
    ExpectOk "trace 开始描迹", "actTraceStart"
    Chk "trace 洋红十字在图形中心", MarkAt(NM_TR, OX, OY, 0.01), "count=" & CountS(NM_TR)
    MoveMark NM_TR, 260!, 200!
    ExpectOk "trace 记一点(手动)", "actTracePoint"
    MoveMark NM_TR, 360!, 260!
    Pause 1.3                                     ' 等 400ms 定时器自动记点
    ExpectOk "trace 完成描迹", "actTraceDone"
    n0 = NEff()
    ExpectOk "trace 生成", "actGenerate"
    Chk "trace 生成新增 1 个动画", NEff() = n0 + 1, "effects " & n0 & " -> " & NEff()
    If NEff() > n0 Then
        n = ReadPath(NEff(), pts, shp, dur, trig)
        Chk "trace 路径点数=3 (起点, 手动点, 定时器自动点)", n = 3, "n=" & n
        If n = 3 Then
            Chk "trace 手动点", (Abs(pts(1)(0) - 100!) < 0.01 And Abs(pts(1)(1) - 60!) < 0.01), "(" & Fmt(pts(1)(0)) & "," & Fmt(pts(1)(1)) & ")"
            Chk "trace 定时器自动记的点", (Abs(pts(2)(0) - 200!) < 0.01 And Abs(pts(2)(1) - 120!) < 0.01), "(" & Fmt(pts(2)(0)) & "," & Fmt(pts(2)(1)) & ")"
        End If
    End If
    Chk "trace 生成后标记已清理", CountS(NM_TR) = 0, "TR=" & CountS(NM_TR)

    ' ===== 10. 过期图形: 移动图形被删后再操作, 应友好提示而不是报错号 =====
    SelMover
    Application.Run "SetForm", "直线"
    ExpectOk "stale 开始", "actStart"
    gS.Shapes("FT_mover").Delete
    SelTarget
    n0 = NEff()
    ExpectOk "stale 图形已删后 放C点", "actPlaceC"
    ExpectOk "stale 图形已删后 生成", "actGenerate"
    ExpectOk "stale 图形已删后 +航点", "actAddW"
    ExpectOk "stale 图形已删后 清空航点", "actClearW"
    Chk "stale 不产生动画", NEff() = n0, "effects " & n0 & " -> " & NEff()
    ExpectOk "stale 清理标记", "actClean"
    Chk "stale 清理后无标记", (gS.Shapes.Count = 1), "shapes=" & gS.Shapes.Count

    ' ===== 11. 简易宏 =====
    SimpleSuite
    GoTo done
eh:
    Chk "driver crashed", False, "ERR " & Err.Number & " " & Err.Description
done:
    W "SUMMARY pass=" & gPass & " fail=" & gFail
End Sub

Private Function NEffOf(ByVal nm As String) As Long
    Dim i As Long
    For i = 1 To NEff()
        If gS.TimeLine.MainSequence(i).Shape.Name = nm Then NEffOf = NEffOf + 1
    Next i
End Function

' 层叠在下的先创建: t2 在下, m2 在上 => ShapeRange(1)=t2
Private Function AddPair()
    Dim t2 As Shape, m2 As Shape
    Set t2 = gS.Shapes.AddShape(msoShapeOval, 100!, 400!, 80!, 60!)
    t2.Name = "FT_t2"                              ' 中心 (140,430)
    Set m2 = gS.Shapes.AddShape(msoShapeRectangle, 500!, 60!, 80!, 60!)
    m2.Name = "FT_m2"                              ' 中心 (540,90)
End Function

Private Sub SimpleSuite()
    Dim n0 As Long, n As Long, pts() As Variant, shp As String, dur As Single, trig As Long
    Dim r As String, m2 As Shape, t2 As Shape
    InitScene
    SetPolicy "ok"
    ' --- M0 没选够图形: 只提示, 不产生动画 ---
    SelNone
    n0 = NEff()
    r = RunM("精确路径_中心到中心")
    Chk "simple 未选图形: 不产生动画(提示框见看门狗日志)", (r = "ok" And NEff() = n0), r & "  effects " & n0 & "->" & NEff()

    ' --- M1 选中 移动图形 + Ctrl 目标; 确认框答"是" ---
    SetPolicy "yes"
    gMover.Select
    gTarget.Select msoFalse
    n0 = NEff()
    r = RunM("精确路径_中心到中心")
    Chk "simple 中心到中心: 无错误", (r <> "VBAERRDIALOG" And Not r Like "ERR*"), r
    If NEff() > n0 Then n = ReadPath(NEff(), pts, shp, dur, trig)
    Chk "simple 中心到中心: FT_mover 移到 FT_target 中心", (NEff() = n0 + 1 And shp = "FT_mover" And EndErr(pts, n, TX - OX, TY - OY) < 0.01), _
        "effects " & n0 & "->" & NEff() & " shape=" & shp & " enderr=" & Fmt(EndErr(pts, n, TX - OX, TY - OY))

    ' --- M2 层叠顺序与点选意图相反的一对: 用户先点 m2 再 Ctrl 点 t2, 意图 m2 -> t2 ---
    AddPair
    Set m2 = FindS("FT_m2"): Set t2 = FindS("FT_t2")
    m2.Select
    t2.Select msoFalse
    SetPolicy "no"                               ' 确认框默认方向是 ShapeRange(1)->(2) 即 t2->m2, 答"否"取反向
    n0 = NEff()
    r = RunM("精确路径_中心到中心")
    shp = "": n = 0
    If NEff() > n0 Then n = ReadPath(NEff(), pts, shp, dur, trig)
    ' ShapeRange 按点选顺序返回(probe E4 已证): 默认方向就是 m2->t2, 答"否"对调为 t2->m2
    Chk "simple 方向可控(答否=对调): 效果落在 FT_t2, 终点=FT_m2 中心", (NEff() = n0 + 1 And shp = "FT_t2" And EndErr(pts, n, 540! - 140!, 90! - 430!) < 0.01), _
        r & "  effects " & n0 & "->" & NEff() & " shape=" & shp

    ' --- M3 确认框答"取消": 不产生动画 ---
    SetPolicy "cancel"
    m2.Select
    t2.Select msoFalse
    n0 = NEff()
    r = RunM("精确路径_中心到中心")
    Chk "simple 确认框取消: 不产生动画", (NEff() = n0 And r <> "VBAERRDIALOG"), r & "  effects " & n0 & "->" & NEff()

    ' --- M4 到坐标(用默认值) ---
    SetPolicy "ok"
    m2.Select
    n0 = NEff()
    r = RunM("精确路径_到坐标")
    shp = "": n = 0
    If NEff() > n0 Then n = ReadPath(NEff(), pts, shp, dur, trig)
    Chk "simple 到坐标: 在所选图形上生成 1 个动画", (NEff() = n0 + 1 And shp = "FT_m2" And n >= 2), r & "  effects " & n0 & "->" & NEff() & " shape=" & shp
    If n >= 2 Then Info "simple 到坐标: 落点=(" & Fmt(540! + pts(n - 1)(0)) & "," & Fmt(90! + pts(n - 1)(1)) & ")"

    ' --- M5 到坐标: InputBox 点取消 ---
    SetPolicy "cancel"
    m2.Select
    n0 = NEff()
    r = RunM("精确路径_到坐标")
    Chk "simple 到坐标取消: 不报错不产生动画", (NEff() = n0 And r <> "VBAERRDIALOG" And Not r Like "ERR*"), r & "  effects " & n0 & "->" & NEff()

    ' --- M6 删除本图形动画: 只删所选图形的 ---
    SetPolicy "ok"
    m2.Select
    r = RunM("删除本图形动画")
    Chk "simple 删除动画: 只删所选图形(FT_m2)的", (NEffOf("FT_m2") = 0 And NEffOf("FT_mover") > 0 And r <> "VBAERRDIALOG"), _
        r & "  m2=" & NEffOf("FT_m2") & " mover=" & NEffOf("FT_mover") & " total=" & NEff()
    SetPolicy "ok"
End Sub

' 面板显示/隐藏压测(回归: 显示后按 UserForms 下标读状态会间歇性"下标越界")
Public Sub FT_PanelStress()
    gPass = 0: gFail = 0
    On Error GoTo eh
    On Error Resume Next
    Kill Environ$("TEMP") & "\abp_ft_results.txt"
    On Error GoTo 0
    InitScene
    Dim k As Long, j As Long, r As String, badK As Long
    Dim d0 As Long: d0 = CountStr(DlgText(), "DIALOG title=[")
    SelMover
    r = RunM("actStart")
    Chk "stress 开始", r = "ok", r
    For k = 1 To 30
        r = RunM("TogglePanel")
        If r <> "ok" Then badK = k: Exit For
        Pause 0.25
        For j = 1 To 40
            r = RunM("PanelLoaded")
            If r <> "ok" Then badK = k: Exit For
        Next j
        If r <> "ok" Then Exit For
        r = RunM("TogglePanel")
        If r <> "ok" Then badK = k: Exit For
        Pause 0.1
    Next k
    Chk "stress 30轮 显示/隐藏 + 1200次读面板状态 无错误", (badK = 0), IIf(badK = 0, "done", "failed at cycle " & badK & " => " & r)
    Chk "stress 全程无弹窗", CountStr(DlgText(), "DIALOG title=[") = d0, "dialogs +" & (CountStr(DlgText(), "DIALOG title=[") - d0)
    r = RunM("TogglePanel")        ' 若面板还开着则收起
    W "SUMMARY pass=" & gPass & " fail=" & gFail
    Exit Sub
eh:
    Chk "stress driver crashed", False, "ERR " & Err.Number & " " & Err.Description
    W "SUMMARY pass=" & gPass & " fail=" & gFail
End Sub

' 未点开始就点"清空航点": 现版本会直接抛 VBA 原始错误并终止本驱动
Public Sub FT_PreStart()
    gPass = 0: gFail = 0
    On Error GoTo eh
    On Error Resume Next
    Kill Environ$("TEMP") & "\abp_ft_results.txt"
    On Error GoTo 0
    InitScene
    SelNone
    Chk "pre-start 清空航点", (RunM("actClearW") = "ok"), RunM("actClearW")
    W "SUMMARY pass=" & gPass & " fail=" & gFail
    Exit Sub
eh:
    Chk "driver crashed", False, "ERR " & Err.Number & " " & Err.Description
    W "SUMMARY pass=" & gPass & " fail=" & gFail
End Sub


' 探针: 无选择时 简易宏 的守卫行为(为何没弹提示框)
Public Sub FT_ProbeSimple()
    gPass = 0: gFail = 0
    On Error Resume Next
    Kill Environ$("TEMP") & "\abp_ft_results.txt"
    On Error GoTo 0
    InitScene
    SelNone
    W "probe selType=" & ActiveWindow.Selection.Type
    Dim sc As Variant
    Err.Clear
    sc = Empty
    sc = ActiveWindow.Selection.ShapeRange.Count
    W "probe srCount=" & sc & " err=" & Err.Number
    SetPolicy "ok"
    Dim r As String
    r = RunM("精确路径_中心到中心")
    W "probe run => " & r
    W "SUMMARY pass=" & gPass & " fail=" & gFail
End Sub

' 调试: 复现 RunM 的弹窗计数, 直接落盘观察
Public Sub FT_RunMDebug()
    InitScene
    SelNone
    SetPolicy "ok"
    Dim fn As String: fn = Environ$("TEMP") & "\abp_ft_results.txt"
    Dim f As Integer
    Dim dlgA As String, d0 As Long, dN As Long
    f = FreeFile
    Open fn For Output As #f
    Print #f, "dbg start"
    Close #f
    dlgA = DlgText()
    d0 = CountStr(dlgA, "DIALOG title=[")
    f = FreeFile
    Open fn For Append As #f
    Print #f, "dbg len=" & Len(dlgA) & " d0=" & d0
    Close #f
    On Error Resume Next
    Application.Run "精确路径_中心到中心"
    On Error GoTo 0
    Pause 0.5
    dlgA = DlgText()
    dN = CountStr(dlgA, "DIALOG title=[") - d0
    f = FreeFile
    Open fn For Append As #f
    Print #f, "dbg len2=" & Len(dlgA) & " dN=" & dN & " head=[" & Left$(dlgA, 60) & "]"
    Close #f
End Sub

' 调试2: 细看 VBA 打开 dialogs.log 的具体错误
Public Sub FT_DlgFileDebug()
    Dim fn As String: fn = Environ$("TEMP") & "\abp_ft_dialogs.log"
    Dim fnr As String: fnr = Environ$("TEMP") & "\abp_ft_results.txt"
    Dim f As Integer
    f = FreeFile
    Open fnr For Output As #f
    Print #f, "dbg2 exists(dialogs)=" & (Len(Dir$(fn)) > 0) & " len=" & Len(Dir$(fn))
    Print #f, "dbg2 exists(results)=" & (Len(Dir$(fnr)) > 0)
    Dim s As String, e1 As Long, d1 As String, lf As Long
    Err.Clear
    On Error Resume Next
    f = FreeFile
    Open fn For Input As #f
    e1 = Err.Number: d1 = Err.Description
    If e1 = 0 Then
        lf = LOF(f)
        s = Input(lf, f)
        Print #f, "dbg2 open ok lof=" & lf & " readlen=" & Len(s)
        Close #f
    Else
        Print #f, "dbg2 open FAIL err=" & e1 & " " & d1
    End If
    Close #f
End Sub

' 调试3: VBA 的 TEMP 到底是哪
Public Sub FT_TempDebug()
    Dim fnr As String: fnr = Environ$("TEMP") & "\abp_ft_results.txt"
    Dim f As Integer
    f = FreeFile
    Open fnr For Append As #f
    Print #f, "dbg3 vbaTEMP=[" & Environ$("TEMP") & "]"
    Print #f, "dbg3 policy exists=" & (Len(Dir$(Environ$("TEMP") & "\abp_ft_policy.txt")) > 0)
    Print #f, "dbg3 dialogs exists=" & (Len(Dir$(Environ$("TEMP") & "\abp_ft_dialogs.log")) > 0)
    Close #f
End Sub

' 调试4: VBA 视角下 dialogs 文件的存在性/Binary 读取
Public Sub FT_FileDebug2()
    Dim fn As String: fn = Environ$("TEMP") & "\abp_ft_dialogs.log"
    Dim fnr As String: fnr = Environ$("TEMP") & "\abp_ft_results.txt"
    Dim f As Integer
    f = FreeFile
    Open fnr For Output As #f
    Print #f, "dbg4 dir(dialogs)=[" & Dir$(fn) & "] len=" & Len(Dir$(fn))
    Print #f, "dbg4 attr=" & GetAttr(fn)
    Close #f
eh:
    f = FreeFile
    Open fnr For Append As #f
    Print #f, "dbg4 err=" & Err.Number & " " & Err.Description
    Close #f
End Sub
