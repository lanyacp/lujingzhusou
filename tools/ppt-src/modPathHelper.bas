Attribute VB_Name = "modPathHelper"
Option Explicit

' =====================================================================
'  路径助手 · PowerPoint 版 (VBA)
'  移植自 WPS 版路径助手: 按图形中心/标记点精确生成路径动画
'  曲线(抛物线/圆弧/航点)用多段直线接力逼近
' =====================================================================

Public Const MK As String = "__pthlp_"
Public Const NM_A As String = "__pthlp_A__"
Public Const NM_B As String = "__pthlp_B__"
Public Const NM_PC As String = "__pthlp_PC__"
Public Const NM_PD As String = "__pthlp_PD__"
Public Const NM_PM As String = "__pthlp_PM__"
Public Const NM_WP As String = "__pthlp_W"
Public Const NM_TR As String = "__pthlp_TR__"

Public gMover As Shape
Public gSlide As Slide
Public gForm As String   ' line/parabola/arc/waypoints/trace

' 描迹状态
Public gTraceOn As Boolean
Public gTraceN As Long
Public gTracePts() As Variant

' ---------------- 面板 ----------------

Public Sub ShowPanel()
    On Error GoTo eh
    frmPanel.StartUpPosition = 1   ' CenterOwner
    frmPanel.Show vbModeless
    Exit Sub
eh:
    MsgBox "面板启动失败: " & Err.Description & vbCrLf & vbCrLf & _
           "若提示宏被禁用: 文件→选项→信任中心→信任中心设置→宏设置" & vbCrLf & _
           "选「禁用所有宏，并发出通知」→确定 → 关闭文件重新打开 → 点黄条「启用内容」", _
           vbExclamation, "路径助手"
End Sub

Public Function PanelLoaded() As Boolean
    Dim i As Long
    For i = 1 To UserForms.Count
        If UserForms(i).Name = "frmPanel" Then PanelLoaded = True: Exit Function
    Next i
End Function

Public Sub TogglePanel()
    If PanelLoaded() Then
        Unload frmPanel
    Else
        ShowPanel
    End If
End Sub

' ---------------- 加载项(.ppam)支持 ----------------
' ppam 加载时自动运行: 建"路径助手"工具条按钮 + 自动弹出面板

Public Sub Auto_Open()
    On Error Resume Next
    BuildBar   ' 只建常驻工具条; 面板用工具条按钮一键开, 跨所有PPT有效
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Dim i As Long
    For i = 1 To UserForms.Count
        Unload UserForms(i)
    Next i
    Application.CommandBars("ABPathBar").Delete
End Sub

Private Sub BuildBar()
    On Error Resume Next
    Application.CommandBars("ABPathBar").Delete
    Dim cb As Object, btn As Object
    Set cb = Application.CommandBars.Add("ABPathBar", 1, False, True)  ' msoBarTop, Temporary
    Set btn = cb.Controls.Add(1)  ' msoControlButton
    btn.Caption = "路径助手面板"
    btn.OnAction = "TogglePanel"
    btn.Style = 2                  ' msoButtonCaption
    cb.Visible = True
End Sub

Public Sub SetForm(nm As String)
    Select Case nm
        Case "抛物线": gForm = "parabola"
        Case "圆弧": gForm = "arc"
        Case "航点": gForm = "waypoints"
        Case "描迹": gForm = "trace"
        Case Else: gForm = "line"
    End Select
    If gForm <> "line" Then
        ' 曲线形态无 B 标记
        Dim b As Shape: Set b = FindMark(gSlide, NM_B)
        If Not b Is Nothing Then b.Delete
    End If
    If gForm <> "trace" Then gTraceOn = False   ' 离开描迹停止轮询
End Sub

' ---------------- 基础工具 ----------------

Private Function isMark(sh As Shape) As Boolean
    isMark = (Left$(sh.Name, 8) = MK)
End Function

Private Function CenterOf(sh As Shape) As Variant
    CenterOf = Array(sh.Left + sh.Width / 2!, sh.Top + sh.Height / 2!)
End Function

Public Function FindMark(s As Slide, nm As String) As Shape
    Dim i As Long, sh As Shape
    If s Is Nothing Then Exit Function
    For i = 1 To s.Shapes.Count
        Set sh = s.Shapes(i)
        If sh.Name = nm Then Set FindMark = sh: Exit Function
    Next i
End Function

Public Function AddCross(s As Slide, nm As String, cx As Single, cy As Single, sz As Single, rgbv As Long) As Shape
    Dim sh As Shape
    Set sh = s.Shapes.AddShape(msoShapeCross, cx - sz / 2!, cy - sz / 2!, sz, sz)
    sh.Name = nm
    sh.Line.Visible = msoFalse
    sh.Fill.ForeColor.RGB = rgbv
    sh.Fill.Transparency = 0.1
    Set AddCross = sh
End Function

' 吸附: 找 12pt 内最近的目标中心(普通图形 或 中心十字), 返回 Variant(x,y); 无则返回原点值
Private Function SnapPt(s As Slide, x As Single, y As Single) As Variant
    Dim best As Variant, bd As Single
    bd = 12!
    Dim i As Long, sh As Shape, c As Variant, d As Single
    For i = 1 To s.Shapes.Count
        Set sh = s.Shapes(i)
        Dim isCross As Boolean
        isCross = (Left$(sh.Name, 10) = MK & "C_")
        If isCross Then
            ' 中心十字是吸附目标
        ElseIf Left$(sh.Name, 8) = MK Then
            GoTo nxt
        End If
        c = CenterOf(sh)
        d = Sqr((c(0) - x) ^ 2 + (c(1) - y) ^ 2)
        If d < bd And d > 0.01 Then bd = d: best = c
nxt:
    Next i
    If IsEmpty(best) Then best = Array(x, y)
    SnapPt = best
End Function

Public Sub st(msg As String)
    On Error Resume Next
    frmPanel.lblStatus.Caption = msg
    frmPanel.Caption = "路径助手 - " & msg
End Sub

Private Function SelOther() As Shape
    ' 当前选中的"另一个图形"(非标记且非移动图形)
    On Error GoTo done
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Function
    Dim sr As ShapeRange: Set sr = ActiveWindow.Selection.ShapeRange
    If sr.Count <> 1 Then Exit Function
    Dim sh As Shape: Set sh = sr(1)
    If isMark(sh) Then Exit Function
    If Not gMover Is Nothing Then
        If sh.Name = gMover.Name Then Exit Function
    End If
    Set SelOther = sh
done:
End Function

' ---------------- 动作 ----------------

Public Sub actCenterCross()
    On Error GoTo eh
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then st "请先选中图形": Exit Sub
    Dim sr As ShapeRange: Set sr = ActiveWindow.Selection.ShapeRange
    Dim sh As Shape, c As Variant, nm As String, ex As Shape, nOn As Long, nOff As Long
    Dim i As Long
    For i = 1 To sr.Count
        Set sh = sr(i)
        If Not isMark(sh) Then
            nm = MK & "C_" & sh.Id
            Set ex = FindMark(sh.Parent, nm)
            If ex Is Nothing Then
                c = CenterOf(sh)
                AddCross sh.Parent, nm, c(0), c(1), 7, RGB(255, 0, 0)
                nOn = nOn + 1
            Else
                ex.Delete
                nOff = nOff + 1
            End If
        End If
    Next i
    If nOff > 0 Then st "已移除 " & nOff & " 个中心十字" Else st "已为 " & nOn & " 个图形加中心十字(吸附目标)"
    Exit Sub
eh: st "中心十字失败: " & Err.Description
End Sub

Public Sub actStart()
    On Error GoTo eh
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then st "请先选中一个图形": Exit Sub
    Dim sr As ShapeRange: Set sr = ActiveWindow.Selection.ShapeRange
    If sr.Count <> 1 Then st "请只选中一个图形": Exit Sub
    Dim sh As Shape: Set sh = sr(1)
    If isMark(sh) Then st "选中的是辅助标记, 请选普通图形": Exit Sub
    Set gMover = sh
    Set gSlide = sh.Parent
    If gForm = "" Then gForm = "line"
    CleanFlow gSlide
    Dim c As Variant: c = CenterOf(sh)
    AddCross gSlide, NM_A, c(0), c(1), 7, RGB(255, 0, 0)
    If gForm = "line" Then AddCross gSlide, NM_B, c(0), c(1), 8, RGB(0, 112, 255)
    st "已开始: 红十字=中心" & IIf(gForm = "line", "; 拖蓝十字到目标再点生成", _
        IIf(gForm = "trace", "; 点「开始描迹」拖洋红十字描曲线", "; 放好标记后点生成"))
    Exit Sub
eh: st "开始失败: " & Err.Description
End Sub

Public Sub actPlaceC()
    PlaceCommon NM_PC, -42!, -32!, RGB(0, 176, 80), "C点(绿)"
End Sub
Public Sub actPlaceD()
    PlaceCommon NM_PD, 48!, 36!, RGB(255, 140, 0), "D点(橙)"
End Sub
Public Sub actPlaceP()
    PlaceCommon NM_PM, 24!, -70!, RGB(128, 0, 192), "P/M点(紫)"
End Sub

Private Sub PlaceCommon(nm As String, ox As Single, oy As Single, rgbv As Long, tag As String)
    On Error GoTo eh
    If gMover Is Nothing Then st "请先点「开始」": Exit Sub
    Dim c As Variant: c = CenterOf(gMover)
    Dim a As Variant: a = c
    Dim pc As Shape: Set pc = FindMark(gSlide, NM_PC)
    If Not pc Is Nothing Then a = CenterOf(pc)
    Dim tx As Single, ty As Single
    tx = a(0) + ox: ty = a(1) + oy
    Dim atShape As String: atShape = ""
    If nm = NM_PD Then
        Dim o As Shape: Set o = SelOther()
        If Not o Is Nothing Then
            Dim oc As Variant: oc = CenterOf(o)
            tx = oc(0): ty = oc(1): atShape = "(在所选图形中心)"
        End If
    End If
    Dim ex As Shape: Set ex = FindMark(gSlide, nm)
    If Not ex Is Nothing Then ex.Delete
    AddCross gSlide, nm, tx, ty, 7, rgbv
    st "已放 " & tag & atShape & ", 可拖动微调"
    Exit Sub
eh: st "放置失败: " & Err.Description
End Sub

Public Sub actAddW()
    On Error GoTo eh
    If gMover Is Nothing Then st "请先点「开始」": Exit Sub
    Dim c As Variant: c = CenterOf(gMover)
    Dim a As Variant: a = c
    Dim pc As Shape: Set pc = FindMark(gSlide, NM_PC)
    If Not pc Is Nothing Then a = CenterOf(pc)
    Dim n As Long: n = CountW() + 1
    Dim tx As Single, ty As Single
    Dim o As Shape: Set o = SelOther()
    If Not o Is Nothing Then
        Dim oc As Variant: oc = CenterOf(o)
        tx = oc(0): ty = oc(1)
    Else
        tx = a(0) + 30! + 22! * (n - 1): ty = a(1) + 28!
    End If
    AddCross gSlide, NM_WP & n & "__", tx, ty, 7, RGB(0, 168, 192)
    st "已加航点 W" & n & IIf(o Is Nothing, "", "(在所选图形中心)")
    Exit Sub
eh: st "加航点失败: " & Err.Description
End Sub

Private Function CountW() As Long
    Dim i As Long, sh As Shape
    If gSlide Is Nothing Then Exit Function
    For i = 1 To gSlide.Shapes.Count
        Set sh = gSlide.Shapes(i)
        If Left$(sh.Name, 10) = (MK & "W") Then CountW = CountW + 1
    Next i
End Function

Public Sub actClearW()
    DeleteByPrefix gSlide, MK & "W"
    st "已清空航点"
End Sub

' ---------------- 描迹 ----------------
' 原理与WPS版一致: 拖动图形时VBA被阻塞, 松手瞬间OnTime恢复并记录该点;
' 「记一点」为手动兜底; 完成后点「完成」停止轮询

Public Sub actTraceStart()
    On Error GoTo eh
    If gMover Is Nothing Then st "请先点「开始」选图形": Exit Sub
    Dim c As Variant: c = CenterOf(gMover)
    Dim anchor As Variant: anchor = c
    Dim pc As Shape: Set pc = FindMark(gSlide, NM_PC)
    If Not pc Is Nothing Then anchor = CenterOf(pc)
    ' 恢复模式: 已有未激活点列则从末点继续
    If (Not gTraceOn) And gTraceN >= 1 Then
        Dim lastP As Variant: lastP = gTracePts(gTraceN)
        AddCross gSlide, NM_TR, lastP(0), lastP(1), 9, RGB(255, 0, 255)
        gTraceOn = True
        ScheduleTick
        st "继续描迹(已" & (gTraceN + 1) & "点): 拖标记松手自动记; 完成点「完成」"
        Exit Sub
    End If
    Dim ex As Shape: Set ex = FindMark(gSlide, NM_TR)
    If Not ex Is Nothing Then ex.Delete
    AddCross gSlide, NM_TR, anchor(0), anchor(1), 9, RGB(255, 0, 255)
    ReDim gTracePts(0 To 200)
    gTracePts(0) = anchor
    gTraceN = 0
    gTraceOn = True
    ScheduleTick
    st "描迹中: 拖洋红十字沿曲线走, 每段松一下手即自动记点(或点「记一点」)"
    Exit Sub
eh: st "描迹启动失败: " & Err.Description
End Sub

Public Sub actTracePoint()
    On Error GoTo eh
    If Not gTraceOn Then st "请先点「开始描迹」": Exit Sub
    Dim tr As Shape: Set tr = FindMark(gSlide, NM_TR)
    If tr Is Nothing Then st "找不到描迹标记": Exit Sub
    Dim c As Variant: c = CenterOf(tr)
    Dim last As Variant: last = gTracePts(gTraceN)
    If Sqr((c(0) - last(0)) ^ 2 + (c(1) - last(1)) ^ 2) < 0.5 Then
        st "标记未移动, 拖到新位置再记"
        Exit Sub
    End If
    PushTracePoint c
    st "已记 " & (gTraceN + 1) & " 点; 继续拖→记; 走完点「完成」"
    Exit Sub
eh: st "记一点失败: " & Err.Description
End Sub

Public Sub actTraceDone()
    On Error Resume Next
    If Not gTraceOn And gTraceN = 0 Then st "尚未描迹": Exit Sub
    gTraceOn = False
    Dim i As Long
    Dim out() As Variant
    ReDim out(0 To gTraceN)
    For i = 0 To gTraceN
        out(i) = gTracePts(i)
    Next i
    out = SmoothMA(out)
    gTraceN = UBound(out)
    For i = 0 To gTraceN
        gTracePts(i) = out(i)
    Next i
    st "描迹完成(" & (gTraceN + 1) & "点), 点「生成路径动画」"
End Sub

' OnTime 轮询: 松手瞬间读标记位置并记录(拖动中VBA阻塞, 恢复即记)
Public Sub TraceTick()
    If Not gTraceOn Then Exit Sub
    On Error Resume Next
    Dim tr As Shape: Set tr = FindMark(gSlide, NM_TR)
    If Not tr Is Nothing Then
        Dim c As Variant: c = CenterOf(tr)
        Dim last As Variant: last = gTracePts(gTraceN)
        If Sqr((c(0) - last(0)) ^ 2 + (c(1) - last(1)) ^ 2) > 2.5 Then
            PushTracePoint c
        End If
    End If
    ScheduleTick
End Sub

Private Sub ScheduleTick()
    On Error Resume Next
    Application.OnTime Now + 0.4 / 86400#, "TraceTick"
End Sub

Private Sub PushTracePoint(c As Variant)
    If gTraceN >= 200 Then Exit Sub
    gTraceN = gTraceN + 1
    gTracePts(gTraceN) = Array(c(0), c(1))
End Sub

' 三点滑动平均轻量平滑(保留首尾)
Private Function SmoothMA(pts As Variant) As Variant
    Dim n As Long: n = UBound(pts)
    If n < 4 Then SmoothMA = pts: Exit Function
    Dim out() As Variant
    ReDim out(0 To n)
    Dim i As Long
    out(0) = pts(0)
    For i = 1 To n - 1
        out(i) = Array((pts(i - 1)(0) + pts(i)(0) + pts(i + 1)(0)) / 3!, _
                       (pts(i - 1)(1) + pts(i)(1) + pts(i + 1)(1)) / 3!)
    Next i
    out(n) = pts(n)
    SmoothMA = out
End Function

Public Sub actClean()
    On Error Resume Next
    Dim s As Slide
    For Each s In ActivePresentation.Slides
        CleanFlow s
        DeleteByPrefix s, MK & "C_"
    Next s
    st "已清理全部辅助标记"
End Sub

' ---------------- 生成 ----------------

Public Sub actGenerate()
    On Error GoTo eh
    If gMover Is Nothing Then st "请先点「开始」": Exit Sub
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim sw As Single, shh As Single
    sw = pres.PageSetup.SlideWidth: shh = pres.PageSetup.SlideHeight
    If gForm = "" Then gForm = "line"

    ' 锚点
    Dim anchor As Variant
    Dim pc As Shape: Set pc = FindMark(gSlide, NM_PC)
    If pc Is Nothing Then anchor = CenterOf(gMover) Else anchor = CenterOf(pc)

    ' 各标记(带吸附)
    Dim pd As Shape, pm As Shape, b As Shape
    Set pd = FindMark(gSlide, NM_PD)
    Set pm = FindMark(gSlide, NM_PM)
    Set b = FindMark(gSlide, NM_B)

    Dim pts As Variant

    If gForm = "line" Then
        Dim S As Variant, E As Variant
        If Not pc Is Nothing And Not pd Is Nothing Then
            S = anchor
            E = SnapPt(gSlide, CenterOf(pd)(0), CenterOf(pd)(1))
        ElseIf Not b Is Nothing Then
            S = CenterOf(gMover)
            E = SnapPt(gSlide, CenterOf(b)(0), CenterOf(b)(1))
        Else
            st "缺少 B 标记: 请拖蓝十字到目标": Exit Sub
        End If
        pts = Array(S, E)
    ElseIf gForm = "parabola" Then
        If pm Is Nothing Or pd Is Nothing Then st "抛物线需要 P(顶点) 和 D(终点)": Exit Sub
        Dim Pp As Variant, Ee As Variant
        Pp = SnapPt(gSlide, CenterOf(pm)(0), CenterOf(pm)(1))
        Ee = SnapPt(gSlide, CenterOf(pd)(0), CenterOf(pd)(1))
        pts = SampleParabola(anchor, Pp, Ee)
    ElseIf gForm = "arc" Then
        If pm Is Nothing Or pd Is Nothing Then st "圆弧需要 M(弧点) 和 D(终点)": Exit Sub
        Dim Mm As Variant
        Mm = SnapPt(gSlide, CenterOf(pm)(0), CenterOf(pm)(1))
        Ee = SnapPt(gSlide, CenterOf(pd)(0), CenterOf(pd)(1))
        pts = SampleArc(anchor, Mm, Ee)
    ElseIf gForm = "trace" Then
        ' 描迹: 用已记录的点列
        If gTraceN < 1 Then st "请先「开始描迹」并拖动标记记录点": Exit Sub
        ReDim pts(0 To gTraceN)
        Dim q As Long
        For q = 0 To gTraceN
            pts(q) = gTracePts(q)
        Next q
        pts = SmoothMA(pts)
    Else ' waypoints
        Dim ws As Variant
        ws = CollectW()
        If IsEmpty(ws) Then st "请先「+航点」": Exit Sub
        Dim all() As Variant
        ReDim all(0 To UBound(ws) + 1)
        all(0) = anchor
        Dim k As Long
        For k = 0 To UBound(ws)
            all(k + 1) = ws(k)
        Next k
        pts = all
        If OptSmooth() Then pts = SmoothSubdiv(pts)
    End If

    ' 去除过密点
    pts = PrunePts(pts, 0.5!)

    If UBound(pts) < 1 Then st "点数不足": Exit Sub
    Dim nSeg As Long: nSeg = UBound(pts)

    ' 替换旧动画
    If OptReplace() Then DeleteShapeEffects gSlide, gMover

    Dim n As Long
    n = BuildEffects(gSlide, gMover, pts, TrigVal(), OptDur())

    If Not OptKeep() Then CleanFlow gSlide
    st "已生成 " & n & " 段动画(" & (UBound(pts) + 1) & "点), 已自动吸附最近中心"
    gMover.Select Replace:=msoTrue
    Exit Sub
eh: st "生成失败: " & Err.Description
End Sub

' 多段直线接力: 每段一个 effect, 首段指定触发, 其余接上一段
Public Function BuildEffects(s As Slide, sh As Shape, pts As Variant, firstTrig As Long, totalDur As Single) As Long
    Dim pres As Presentation: Set pres = s.Parent
    Dim sw As Single, shh As Single
    sw = pres.PageSetup.SlideWidth: shh = pres.PageSetup.SlideHeight
    Dim n As Long: n = UBound(pts)
    Dim dur As Single: dur = totalDur / n
    Dim i As Long, dx As Single, dy As Single, eff As Effect, beh As AnimationBehavior, trig As Long
    For i = 1 To n
        If i = 1 Then trig = firstTrig Else trig = msoAnimTriggerAfterPrevious
        dx = (pts(i)(0) - pts(i - 1)(0)) / sw
        dy = (pts(i)(1) - pts(i - 1)(1)) / shh
        Set eff = s.TimeLine.MainSequence.AddEffect(sh, msoAnimEffectCustom, , trig)
        Set beh = eff.Behaviors.Add(msoAnimTypeMotion)
        beh.MotionEffect.FromX = 0: beh.MotionEffect.FromY = 0
        beh.MotionEffect.ToX = dx: beh.MotionEffect.ToY = dy
        eff.Timing.Duration = dur
    Next i
    BuildEffects = n
End Function

' ---------------- 选项读取(容错) ----------------

Private Function OptDur() As Single
    On Error Resume Next
    OptDur = CSng(frmPanel.txtDur.Text)
    If OptDur <= 0 Then OptDur = 2!
End Function
Private Function OptSmooth() As Boolean
    On Error Resume Next
    OptSmooth = frmPanel.chkSmooth.Value
End Function
Private Function OptKeep() As Boolean
    On Error Resume Next
    OptKeep = frmPanel.chkKeep.Value
End Function
Private Function OptReplace() As Boolean
    On Error Resume Next
    OptReplace = frmPanel.chkReplace.Value
End Function
Private Function TrigVal() As Long
    On Error Resume Next
    Dim v As String: v = frmPanel.cmbTrig.Value
    If v = "与上一动画同时" Then
        TrigVal = msoAnimTriggerWithPrevious
    ElseIf v = "上一动画之后" Then
        TrigVal = msoAnimTriggerAfterPrevious
    Else
        TrigVal = msoAnimTriggerOnPageClick
    End If
End Function

' ---------------- 数学: 采样 ----------------

Private Function SampleParabola(S As Variant, P As Variant, E As Variant) As Variant
    ' 曲线精确过 P(顶点): 二次贝塞尔→三次
    Dim q(1) As Single, c1(1) As Single, c2(1) As Single
    q(0) = (4! * P(0) - S(0) - E(0)) / 2!: q(1) = (4! * P(1) - S(1) - E(1)) / 2!
    c1(0) = S(0) + (q(0) - S(0)) * 2! / 3!: c1(1) = S(1) + (q(1) - S(1)) * 2! / 3!
    c2(0) = E(0) + (q(0) - E(0)) * 2! / 3!: c2(1) = E(1) + (q(1) - E(1)) * 2! / 3!
    Dim out() As Variant
    ReDim out(0 To 12)
    Dim i As Long, t As Single, u As Single
    For i = 0 To 12
        t = CSng(i) / 12!: u = 1! - t
        out(i) = Array(u * u * u * S(0) + 3! * u * u * t * c1(0) + 3! * u * t * t * c2(0) + t * t * t * E(0), _
                       u * u * u * S(1) + 3! * u * u * t * c1(1) + 3! * u * t * t * c2(1) + t * t * t * E(1))
    Next i
    SampleParabola = out
End Function

Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then Atan2 = Atn(y / x) + 3.14159265358979 Else Atan2 = Atn(y / x) - 3.14159265358979
    Else
        If y > 0 Then Atan2 = 1.5707963267949 ElseIf y < 0 Then Atan2 = -1.5707963267949 Else Atan2 = 0
    End If
End Function

Private Function NormA(ByVal a As Double) As Double
    Const PI2 As Double = 6.28318530717959
    Do While a < 0: a = a + PI2: Loop
    Do While a >= PI2: a = a - PI2: Loop
    NormA = a
End Function

Private Function SampleArc(S As Variant, M As Variant, E As Variant) As Variant
    Dim d As Double
    d = 2 * (S(0) * (M(1) - E(1)) + M(0) * (E(1) - S(1)) + E(0) * (S(1) - M(1)))
    If Abs(d) < 0.000001 Then
        SampleArc = Array(S, E)
        Exit Function
    End If
    Dim ss As Double, mm As Double, ee As Double
    ss = S(0) * S(0) + S(1) * S(1): mm = M(0) * M(0) + M(1) * M(1): ee = E(0) * E(0) + E(1) * E(1)
    Dim ux As Double, uy As Double
    ux = (ss * (M(1) - E(1)) + mm * (E(1) - S(1)) + ee * (S(1) - M(1))) / d
    uy = (ss * (E(0) - M(0)) + mm * (S(0) - E(0)) + ee * (M(0) - S(0))) / d
    Dim r As Double
    r = Sqr((S(0) - ux) ^ 2 + (S(1) - uy) ^ 2)
    Dim a0 As Double, a1 As Double, am As Double
    a0 = Atan2(S(1) - uy, S(0) - ux)
    a1 = Atan2(E(1) - uy, E(0) - ux)
    am = Atan2(M(1) - uy, M(0) - ux)
    Dim d1 As Double, dm As Double, total As Double
    d1 = NormA(a1 - a0): dm = NormA(am - a0)
    If dm < d1 Then total = d1 Else total = d1 - 6.28318530717959
    Dim n As Long: n = 24
    Dim out() As Variant
    ReDim out(0 To n)
    Dim i As Long, a As Double
    For i = 0 To n
        a = a0 + total * i / n
        out(i) = Array(ux + r * Cos(a), uy + r * Sin(a))
    Next i
    SampleArc = out
End Function

' Catmull-Rom 上采样(每段4点), 让航点平滑
Private Function SmoothSubdiv(pts As Variant) As Variant
    If UBound(pts) < 2 Then SmoothSubdiv = pts: Exit Function
    Dim per As Long: per = 4
    Dim out() As Variant
    ReDim out(0 To (UBound(pts) - 1) * per)
    Dim idx As Long
    idx = 0
    out(0) = pts(0)
    idx = 1
    Dim i As Long, j As Long, t As Single
    Dim p0 As Variant, p1 As Variant, p2 As Variant, p3 As Variant
    For i = 0 To UBound(pts) - 2
        p1 = pts(i): p2 = pts(i + 1)
        p0 = pts(IIf(i = 0, 0, i - 1))
        p3 = pts(IIf(i + 2 > UBound(pts), UBound(pts), i + 2))
        For j = 1 To per
            t = CSng(j) / CSng(per)
            out(idx) = Array(CR1(p0(0), p1(0), p2(0), p3(0), t), CR1(p0(1), p1(1), p2(1), p3(1), t))
            idx = idx + 1
        Next j
    Next i
    ReDim Preserve out(0 To idx - 1)
    SmoothSubdiv = out
End Function

Private Function CR1(a As Single, b As Single, c As Single, d As Single, t As Single) As Single
    Dim t2 As Single, t3 As Single
    t2 = t * t: t3 = t2 * t
    CR1 = 0.5! * ((2! * b) + (-a + c) * t + (2! * a - 5! * b + 4! * c - d) * t2 + (-a + 3! * b - 3! * c + d) * t3)
End Function

Private Function PrunePts(pts As Variant, minD As Single) As Variant
    Dim out() As Variant
    ReDim out(0 To UBound(pts))
    Dim idx As Long: idx = 0
    out(0) = pts(0)
    Dim i As Long
    For i = 1 To UBound(pts)
        If Sqr((pts(i)(0) - out(idx)(0)) ^ 2 + (pts(i)(1) - out(idx)(1)) ^ 2) >= minD Then
            idx = idx + 1
            out(idx) = pts(i)
        End If
    Next i
    ReDim Preserve out(0 To idx)
    PrunePts = out
End Function

Private Function CollectW() As Variant
    Dim tmp() As Variant
    ReDim tmp(0 To 63)
    Dim n As Long: n = 0
    Dim i As Long, sh As Shape, nm As String, j As Long
    ' 按编号排序收集
    Dim idxs(64) As Long, names(64) As String, cnt As Long
    For i = 1 To gSlide.Shapes.Count
        Set sh = gSlide.Shapes(i)
        nm = sh.Name
        If Left$(nm, 10) = (MK & "W") Then
            cnt = cnt + 1
            idxs(cnt) = CLng(Val(Mid$(nm, 11)))
            Set shsTmp(cnt) = sh
        End If
    Next i
    ' 冒泡
    Dim t As Long, ts As String
    For i = 1 To cnt - 1
        For j = 1 To cnt - i
            If idxs(j) > idxs(j + 1) Then
                t = idxs(j): idxs(j) = idxs(j + 1): idxs(j + 1) = t
                Set tsSh = shsTmp(j): Set shsTmp(j) = shsTmp(j + 1): Set shsTmp(j + 1) = tsSh
            End If
        Next j
    Next i
    Dim out() As Variant
    If cnt = 0 Then
        CollectW = Empty
        Exit Function
    End If
    ReDim out(0 To cnt - 1)
    For i = 1 To cnt
        out(i - 1) = SnapPt(gSlide, CenterOf(shsTmp(i))(0), CenterOf(shsTmp(i))(1))
    Next i
    CollectW = out
End Function

' 模块级暂存(排序用)
Private shsTmp(64) As Shape
Private tsSh As Shape

' ---------------- 清理 ----------------

Public Sub CleanFlow(s As Slide)
    On Error Resume Next
    Dim nm As String
    nm = NM_A: KillMark s, nm
    nm = NM_B: KillMark s, nm
    nm = NM_PC: KillMark s, nm
    nm = NM_PD: KillMark s, nm
    nm = NM_PM: KillMark s, nm
    nm = NM_TR: KillMark s, nm
    DeleteByPrefix s, MK & "W"
End Sub

Private Sub KillMark(s As Slide, nm As String)
    On Error Resume Next
    Dim sh As Shape: Set sh = FindMark(s, nm)
    If Not sh Is Nothing Then sh.Delete
End Sub

Public Sub DeleteByPrefix(s As Slide, pfx As String)
    On Error Resume Next
    Dim i As Long
    For i = s.Shapes.Count To 1 Step -1
        If Left$(s.Shapes(i).Name, Len(pfx)) = pfx Then s.Shapes(i).Delete
    Next i
End Sub

Public Sub DeleteShapeEffects(s As Slide, sh As Shape)
    On Error Resume Next
    Dim seq As Sequence: Set seq = s.TimeLine.MainSequence
    Dim i As Long
    For i = seq.Count To 1 Step -1
        If seq(i).Shape.Name = sh.Name Then seq(i).Delete
    Next i
End Sub

' ---------------- 自测(写结果文件) ----------------

Public Sub SelfTest()
    Dim res As String
    On Error GoTo eh
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim s As Slide: Set s = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)
    Dim sh As Shape: Set sh = s.Shapes.AddShape(msoShapeRectangle, 100!, 100!, 120!, 80!)
    sh.Name = "abp_test_rect"
    Set gMover = sh: Set gSlide = s: gForm = "line"
    AddCross s, NM_A, 160!, 140!, 7, RGB(255, 0, 0)
    AddCross s, NM_B, 460!, 280!, 8, RGB(0, 112, 255)
    Dim pts As Variant
    pts = Array(CenterOf(sh), Array(460!, 280!))
    Dim n As Long
    n = BuildEffects(s, sh, pts, msoAnimTriggerOnPageClick, 2!)
    Dim seq As Sequence: Set seq = s.TimeLine.MainSequence
    Dim eff As Effect: Set eff = seq(seq.Count)
    Dim tox As Single, toy As Single
    Dim i As Long
    For i = 1 To eff.Behaviors.Count
        If eff.Behaviors(i).Type = msoAnimTypeMotion Then
            tox = eff.Behaviors(i).MotionEffect.ToX
            toy = eff.Behaviors(i).MotionEffect.ToY
        End If
    Next i
    Dim expX As Single, expY As Single
    expX = 300! / pres.PageSetup.SlideWidth
    expY = 140! / pres.PageSetup.SlideHeight
    If n = 1 And Abs(tox - expX) < 0.0001 And Abs(toy - expY) < 0.0001 Then
        res = "PASS n=" & n & " ToX=" & tox & " ToY=" & toy
    Else
        res = "FAIL n=" & n & " ToX=" & tox & "/" & expX & " ToY=" & toy & "/" & expY
    End If
    s.Delete
    GoTo done
eh:
    res = "ERR " & Err.Number & " " & Err.Description
done:
    Dim f As Integer
    f = FreeFile
    Open Environ$("TEMP") & "\abp_ppt_selftest.txt" For Output As #f
    Print #f, res
    Close #f
End Sub
