Attribute VB_Name = "modFTDriver"
Option Explicit

' 能力探针: 确认 PowerPoint 对象模型在本机的真实行为(产品代码依赖这些行为)

Private f As Integer

Private Sub W(ByVal s As String)
    Print #f, s
End Sub

Private Sub OpenLog()
    f = FreeFile
    Open Environ$("TEMP") & "\abp_ft_results.txt" For Output As #f
End Sub

Private Sub Pause(ByVal sec As Single)
    Dim t As Single: t = Timer
    Do While Timer < t + sec
        DoEvents
    Loop
End Sub

' E1-E4: 路径字符串可写性 / 默认时序 / 选择顺序
Public Sub FT_Probe()
    OpenLog
    On Error GoTo eh
    Dim pres As Presentation: Set pres = ActivePresentation
    W "ENV version=" & Application.Version & " build=" & Application.Build & " arch=" & Environ$("PROCESSOR_ARCHITECTURE")
    Dim s As Slide: Set s = pres.Slides.Add(1, ppLayoutBlank)
    W "SLIDE SW=" & pres.PageSetup.SlideWidth & " SH=" & pres.PageSetup.SlideHeight
    Dim a As Shape, b As Shape
    Set a = s.Shapes.AddShape(msoShapeRectangle, 100!, 100!, 120!, 80!)
    Set b = s.Shapes.AddShape(msoShapeOval, 500!, 300!, 100!, 100!)
    ' --- E4 选择顺序 vs ShapeRange 顺序 ---
    b.Select
    a.Select msoFalse
    W "E4 select b-then-a: (1)=" & ActiveWindow.Selection.ShapeRange(1).Name & " (2)=" & ActiveWindow.Selection.ShapeRange(2).Name
    a.Select
    b.Select msoFalse
    W "E4 select a-then-b: (1)=" & ActiveWindow.Selection.ShapeRange(1).Name & " (2)=" & ActiveWindow.Selection.ShapeRange(2).Name
    ' --- E2 新建自定义动画的默认时序 ---
    Dim seq As Sequence: Set seq = s.TimeLine.MainSequence
    Dim eff As Effect, beh As AnimationBehavior
    Set eff = seq.AddEffect(a, msoAnimEffectCustom, , msoAnimTriggerOnPageClick)
    W "E2 after AddEffect: SmoothStart=" & eff.Timing.SmoothStart & " SmoothEnd=" & eff.Timing.SmoothEnd & " Accel=" & eff.Timing.Accelerate & " Decel=" & eff.Timing.Decelerate & " Dur=" & eff.Timing.Duration & " Trig=" & eff.Timing.TriggerType
    Set beh = eff.Behaviors.Add(msoAnimTypeMotion)
    W "E2 after motion beh: SmoothStart=" & eff.Timing.SmoothStart & " SmoothEnd=" & eff.Timing.SmoothEnd & " Accel=" & eff.Timing.Accelerate & " Decel=" & eff.Timing.Decelerate & " Dur=" & eff.Timing.Duration
    beh.MotionEffect.FromX = 0: beh.MotionEffect.FromY = 0
    beh.MotionEffect.ToX = 0.1: beh.MotionEffect.ToY = 0.2
    W "E1 after From/To: Path=[" & beh.MotionEffect.Path & "]"
    ' --- E1 Path 是否可写 ---
    Dim e1 As Long, d1 As String
    On Error Resume Next
    Err.Clear
    beh.MotionEffect.Path = "M 0 0 L 0.1 0.2 L 0.2 0.1 L 0.3 0.3 E"
    e1 = Err.Number: d1 = Err.Description
    On Error GoTo eh
    W "E1 Path write (same beh): err=" & e1 & " " & d1 & " => Path=[" & beh.MotionEffect.Path & "] ToX=" & beh.MotionEffect.ToX & " ToY=" & beh.MotionEffect.ToY
    Set eff = seq.AddEffect(a, msoAnimEffectCustom, , msoAnimTriggerAfterPrevious)
    Set beh = eff.Behaviors.Add(msoAnimTypeMotion)
    On Error Resume Next
    Err.Clear
    beh.MotionEffect.Path = "M 0 0 L 0.05 0.1 L 0.1 0.2 L 0.1 0.3 E"
    e1 = Err.Number: d1 = Err.Description
    On Error GoTo eh
    W "E1b Path write (fresh beh): err=" & e1 & " " & d1 & " => Path=[" & beh.MotionEffect.Path & "]"
    ' --- E3 From<>0 时存储成什么 ---
    Set eff = seq.AddEffect(a, msoAnimEffectCustom, , msoAnimTriggerAfterPrevious)
    Set beh = eff.Behaviors.Add(msoAnimTypeMotion)
    beh.MotionEffect.FromX = 0.1: beh.MotionEffect.FromY = 0.3
    beh.MotionEffect.ToX = 0.2: beh.MotionEffect.ToY = 0.35
    W "E3 From<>0: Path=[" & beh.MotionEffect.Path & "]"
    W "E3b effects=" & seq.Count & " dur(last)=" & seq(seq.Count).Timing.Duration
    ' --- 保存副本以检查 XML ---
    On Error GoTo eh
    pres.SaveCopyAs Environ$("TEMP") & "\abp_probe.pptx", ppSaveAsOpenXMLPresentation
    W "SAVED " & Environ$("TEMP") & "\abp_probe.pptx"
    W "DONE"
    GoTo done
eh:
    W "ERR " & Err.Number & " " & Err.Description
done:
    Close #f
End Sub

' 复现"面板显示后弹 运行时错误9": 显示面板后泵消息 2 秒
Public Sub FT_ShowErr()
    OpenLog
    On Error Resume Next
    Dim e1 As Long, d1 As String
    Err.Clear
    Application.Run ("ShowPanel")
    e1 = Err.Number: d1 = Err.Description
    W "ShowPanel err=" & e1 & " " & d1
    Pause 2
    W "after 2s pump (reached only if no error dialog)"
    Close #f
End Sub

' 按首次复现到"错误9"那次的完全相同步骤重放: 选区/建动画/PanelLoaded/TogglePanel/泵消息
Public Sub FT_ShowErr2()
    OpenLog
    On Error Resume Next
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim s As Slide: Set s = pres.Slides.Add(1, ppLayoutBlank)
    Dim a As Shape, b As Shape
    Set a = s.Shapes.AddShape(msoShapeRectangle, 100!, 100!, 120!, 80!)
    Set b = s.Shapes.AddShape(msoShapeOval, 500!, 300!, 100!, 100!)
    b.Select
    a.Select msoFalse
    W "sel (1)=" & ActiveWindow.Selection.ShapeRange(1).Name
    Dim seq As Sequence: Set seq = s.TimeLine.MainSequence
    Dim eff As Effect, beh As AnimationBehavior, i As Long
    For i = 1 To 3
        Set eff = seq.AddEffect(a, msoAnimEffectCustom, , msoAnimTriggerAfterPrevious)
        Set beh = eff.Behaviors.Add(msoAnimTypeMotion)
        beh.MotionEffect.Path = "M 0 0 L 0.05 0.1 L 0.1 0.2 L 0.1 0.3 E"
    Next i
    W "effects=" & seq.Count
    Dim v As Variant, e1 As Long, d1 As String
    Err.Clear
    v = Empty
    v = Application.Run("PanelLoaded")
    e1 = Err.Number: d1 = Err.Description
    W "Run(PanelLoaded) err=" & e1 & " v=[" & v & "]"
    Err.Clear
    v = Empty
    v = Application.Run("TogglePanel")
    e1 = Err.Number: d1 = Err.Description
    W "Run(TogglePanel) err=" & e1 & " " & d1
    W "pumping..."
    Pause 0.8
    W "after pump (no error dialog)"
    Close #f
End Sub

' 显示面板后持续泵消息, 由编排器模拟真人点击面板窗口(激活), 观察是否弹运行时错误
Public Sub FT_ShowErr3()
    OpenLog
    On Error Resume Next
    Application.Run ("ShowPanel")
    W "shown; pumping 6s"
    Pause 6
    W "after pump (no error dialog)"
    Close #f
End Sub

' 逐语句重放首次复现到错误9 的那次驱动(原 FT_Probe)
Public Sub FT_ShowErr4()
    OpenLog
    On Error GoTo eh
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim s As Slide: Set s = pres.Slides.Add(1, ppLayoutBlank)
    Dim a As Shape, b As Shape
    Set a = s.Shapes.AddShape(msoShapeRectangle, 100!, 100!, 120!, 80!)
    Set b = s.Shapes.AddShape(msoShapeOval, 500!, 300!, 100!, 100!)
    b.Select
    a.Select msoFalse
    W "E4 " & ActiveWindow.Selection.ShapeRange(1).Name
    a.Select
    b.Select msoFalse
    W "E4b " & ActiveWindow.Selection.ShapeRange(1).Name
    Dim seq As Sequence: Set seq = s.TimeLine.MainSequence
    Dim eff As Effect, beh As AnimationBehavior
    Set eff = seq.AddEffect(a, msoAnimEffectCustom, , msoAnimTriggerOnPageClick)
    W "E2 " & eff.Timing.SmoothStart & eff.Timing.SmoothEnd
    Set beh = eff.Behaviors.Add(msoAnimTypeMotion)
    beh.MotionEffect.FromX = 0: beh.MotionEffect.FromY = 0
    beh.MotionEffect.ToX = 0.1: beh.MotionEffect.ToY = 0.2
    W "E1 " & beh.MotionEffect.Path
    Dim e1 As Long, d1 As String
    On Error Resume Next
    Err.Clear
    beh.MotionEffect.Path = "M 0 0 L 0.1 0.2 L 0.2 0.1 L 0.3 0.3 E"
    e1 = Err.Number: d1 = Err.Description
    On Error GoTo eh
    W "E1w " & e1
    Set eff = seq.AddEffect(a, msoAnimEffectCustom, , msoAnimTriggerAfterPrevious)
    Set beh = eff.Behaviors.Add(msoAnimTypeMotion)
    On Error Resume Next
    Err.Clear
    beh.MotionEffect.Path = "M 0 0 L 0.05 0.1 L 0.1 0.2 L 0.1 0.3 E"
    e1 = Err.Number: d1 = Err.Description
    On Error GoTo eh
    W "E1b " & e1
    Set eff = seq.AddEffect(a, msoAnimEffectCustom, , msoAnimTriggerAfterPrevious)
    Set beh = eff.Behaviors.Add(msoAnimTypeMotion)
    beh.MotionEffect.FromX = 0.1: beh.MotionEffect.FromY = 0.3
    beh.MotionEffect.ToX = 0.2: beh.MotionEffect.ToY = 0.35
    W "E3 " & beh.MotionEffect.Path
    Dim v As Variant
    On Error Resume Next
    Err.Clear
    v = Empty
    v = Application.Run("PanelLoaded")
    e1 = Err.Number: d1 = Err.Description
    On Error GoTo eh
    W "E6 " & e1 & " v=[" & v & "] " & TypeName(v)
    On Error Resume Next
    Err.Clear
    v = Empty
    v = Application.Run("TogglePanel")
    e1 = Err.Number: d1 = Err.Description
    On Error GoTo eh
    W "E6b " & e1
    Pause 0.8
    W "E6c-past-pause"
    Close #f
    Exit Sub
eh:
    W "ERR " & Err.Number & " " & Err.Description
    Close #f
End Sub

' 放大复现概率: 同一前置步骤下, 反复 显示/关闭 面板 N 轮
Public Sub FT_Cycle()
    OpenLog
    On Error Resume Next
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim s As Slide: Set s = pres.Slides.Add(1, ppLayoutBlank)
    Dim a As Shape, b As Shape
    Set a = s.Shapes.AddShape(msoShapeRectangle, 100!, 100!, 120!, 80!)
    Set b = s.Shapes.AddShape(msoShapeOval, 500!, 300!, 100!, 100!)
    b.Select
    a.Select msoFalse
    Dim seq As Sequence: Set seq = s.TimeLine.MainSequence
    Dim eff As Effect, beh As AnimationBehavior, i As Long, k As Long
    Set eff = seq.AddEffect(a, msoAnimEffectCustom, , msoAnimTriggerOnPageClick)
    Set beh = eff.Behaviors.Add(msoAnimTypeMotion)
    beh.MotionEffect.FromX = 0: beh.MotionEffect.FromY = 0
    beh.MotionEffect.ToX = 0.1: beh.MotionEffect.ToY = 0.2
    Set eff = seq.AddEffect(a, msoAnimEffectCustom, , msoAnimTriggerAfterPrevious)
    Set beh = eff.Behaviors.Add(msoAnimTypeMotion)
    beh.MotionEffect.FromX = 0.1: beh.MotionEffect.FromY = 0.3
    beh.MotionEffect.ToX = 0.2: beh.MotionEffect.ToY = 0.35
    Dim v As Variant, e1 As Long
    For k = 1 To 10
        v = Empty
        v = Application.Run("PanelLoaded")
        v = Empty
        v = Application.Run("TogglePanel")
        e1 = Err.Number
        W "cycle " & k & " show err=" & e1
        Pause 0.8
        v = Empty
        v = Application.Run("TogglePanel")
        W "cycle " & k & " hide err=" & Err.Number
        Pause 0.3
    Next k
    W "all cycles done"
    Close #f
End Sub

' 判别实验: 面板显示后只反复读状态(PanelLoaded, 不含 Unload), 看 错误9 是否出在这里
Public Sub FT_Cycle2()
    OpenLog
    On Error Resume Next
    Dim v As Variant, k As Long, j As Long
    For k = 1 To 12
        v = Empty
        v = Application.Run("TogglePanel")        ' show
        W "c" & k & " show err=" & Err.Number
        Pause 0.8
        For j = 1 To 300
            v = Empty
            v = Application.Run("PanelLoaded")    ' 只读状态
        Next j
        W "c" & k & " polled PanelLoaded x300"
        Pause 0.3
        v = Empty
        v = Application.Run("TogglePanel")        ' hide
        W "c" & k & " hide err=" & Err.Number
        Pause 0.3
    Next k
    W "all done"
    Close #f
End Sub

' 看门狗自检: 故意抛运行时错误
Public Sub FT_Bang()
    OpenLog
    W "before bang"
    Dim x As Long
    x = 1 \ 0
    W "after bang (should not be written)"
    Close #f
End Sub

' 显示面板并停留, 供外部截图 / 读取控件
Public Sub FT_PanelHold()
    OpenLog
    On Error Resume Next
    Application.Run ("ShowPanel")
    W "shown"
    Pause 16
    W "hold end"
    Close #f
End Sub

' 真人式操作复现: 外部脚本用鼠标点面板按钮, 本过程只负责布景/选图形/事后核对
Public Sub FT_UiClick()
    OpenLog
    On Error Resume Next
    Dim pres As Presentation: Set pres = ActivePresentation
    Do While pres.Slides.Count > 0
        pres.Slides(1).Delete
    Loop
    Dim s As Slide: Set s = pres.Slides.Add(1, ppLayoutBlank)
    Dim mv As Shape, tg As Shape
    Set mv = s.Shapes.AddShape(msoShapeRectangle, 100!, 100!, 120!, 80!)
    mv.Name = "FT_mover"
    Set tg = s.Shapes.AddShape(msoShapeOval, 500!, 300!, 100!, 100!)
    tg.Name = "FT_target"
    mv.Select
    W "before: selType=" & ActiveWindow.Selection.Type
    Application.Run "ShowPanel"
    W "t0 panel shown; selType=" & ActiveWindow.Selection.Type
    Pause 6                                   ' 期间外部点击: 开始(约2.8s) 放C点(约4.8s)
    W "t6 after C: selType=" & ActiveWindow.Selection.Type & " shapes=" & s.Shapes.Count & _
      " A=" & Cnt(s, "__pthlp_A__") & " B=" & Cnt(s, "__pthlp_B__") & " C=" & Cnt(s, "__pthlp_PC__") & " D=" & Cnt(s, "__pthlp_PD__")
    tg.Select
    W "t6 target selected: selType=" & ActiveWindow.Selection.Type
    Pause 4                                   ' 期间外部点击: 放D点(约1.8s后)
    W "t10 after D: shapes=" & s.Shapes.Count & " C=" & Cnt(s, "__pthlp_PC__") & " D=" & Cnt(s, "__pthlp_PD__")
    Dim sh As Shape
    For Each sh In s.Shapes
        If Left$(sh.Name, 8) = "__pthlp_" Then W "   mark " & sh.Name & " at " & Round(sh.Left + sh.Width / 2, 1) & "," & Round(sh.Top + sh.Height / 2, 1)
    Next sh
    Pause 2
    Close #f
End Sub

Private Function Cnt(ByVal s As Slide, ByVal nm As String) As Long
    Dim i As Long
    For i = 1 To s.Shapes.Count
        If s.Shapes(i).Name = nm Then Cnt = Cnt + 1
    Next i
End Function

' 复现"常驻版 + 免安装版同时加载": 驱动注入到 路径助手.pptm 自己的工程里
Public Sub FT_ToolFlow()
    OpenLog
    On Error Resume Next
    Dim pres As Presentation: Set pres = ActivePresentation
    W "active pres=[" & pres.Name & "] slides=" & pres.Slides.Count & " myUserForms=" & UserForms.Count
    Dim s As Slide: Set s = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)
    Dim mv As Shape, tg As Shape
    Set mv = s.Shapes.AddShape(msoShapeRectangle, 100!, 100!, 120!, 80!)
    mv.Name = "FT_mover"
    Set tg = s.Shapes.AddShape(msoShapeOval, 500!, 300!, 100!, 100!)
    tg.Name = "FT_target"
    Dim v As Variant, e1 As Long
    v = Empty
    v = Application.Run("ShowPanel")            ' 封面按钮就是这个动作
    e1 = Err.Number
    W "after Run ShowPanel: err=" & e1 & " myUserForms=" & UserForms.Count
    mv.Select
    Pause 6                                      ' 外部点击: 开始, 放C点
    W "t6 shapes=" & s.Shapes.Count & " A=" & Cnt(s, "__pthlp_A__") & " B=" & Cnt(s, "__pthlp_B__") & " C=" & Cnt(s, "__pthlp_PC__") & " D=" & Cnt(s, "__pthlp_PD__")
    tg.Select
    Pause 4                                      ' 外部点击: 放D点
    W "t10 shapes=" & s.Shapes.Count & " C=" & Cnt(s, "__pthlp_PC__") & " D=" & Cnt(s, "__pthlp_PD__")
    Pause 2
    Close #f
End Sub

' 两份同时加载: 带选区状态记录
Public Sub FT_ToolFlow2()
    OpenLog
    On Error Resume Next
    Dim pres As Presentation: Set pres = ActivePresentation
    W "active pres=[" & pres.Name & "] presWindows=" & pres.Windows.Count & " appWindows=" & Application.Windows.Count
    Dim s As Slide: Set s = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)
    Dim mv As Shape, tg As Shape
    Set mv = s.Shapes.AddShape(msoShapeRectangle, 100!, 100!, 120!, 80!)
    mv.Name = "FT_mover"
    Set tg = s.Shapes.AddShape(msoShapeOval, 500!, 300!, 100!, 100!)
    tg.Name = "FT_target"
    ActiveWindow.View.GotoSlide s.SlideIndex        ' 切到工作幻灯片(否则选不中)
    Dim v As Variant
    v = Application.Run("ShowPanel")
    W "after Run ShowPanel err=" & Err.Number
    mv.Select
    W "after select: selType=" & ActiveWindow.Selection.Type & " count=" & ActiveWindow.Selection.ShapeRange.Count & " slideIdx=" & ActiveWindow.View.Slide.SlideIndex & " shapesOnSlide=" & s.Shapes.Count
    Pause 6
    W "t6 selType=" & ActiveWindow.Selection.Type & " slideIdx=" & ActiveWindow.View.Slide.SlideIndex & " A=" & Cnt(s, "__pthlp_A__") & " B=" & Cnt(s, "__pthlp_B__") & " C=" & Cnt(s, "__pthlp_PC__")
    tg.Select
    Pause 5
    W "t11 selType=" & ActiveWindow.Selection.Type & " C=" & Cnt(s, "__pthlp_PC__") & " D=" & Cnt(s, "__pthlp_PD__")
    Pause 2
    Close #f
End Sub

' 用户场景回归: 新会话没点「开始」, 选中图形直接点「放C点」→ 自动认定
Public Sub FT_AutoAdopt()
    OpenLog
    On Error Resume Next
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim s As Slide: Set s = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)
    ActiveWindow.View.GotoSlide s.SlideIndex
    Dim mv As Shape, tg As Shape
    Set mv = s.Shapes.AddShape(msoShapeRectangle, 100!, 100!, 120!, 80!)
    mv.Name = "FT_mover"
    Set tg = s.Shapes.AddShape(msoShapeOval, 500!, 300!, 100!, 100!)
    tg.Name = "FT_target"
    Application.Run "ShowPanel"
    Pause 1
    mv.Select
    W "selType=" & ActiveWindow.Selection.Type
    Pause 8                                   ' 外部点击面板上的「放C点」
    W "after UI 放C点: C=" & Cnt(s, "__pthlp_PC__") & " A=" & Cnt(s, "__pthlp_A__") & " B=" & Cnt(s, "__pthlp_B__")
    tg.Select
    On Error Resume Next
    Application.Run "actPlaceD"               ' 宏路径放D(选中目标=自动落其中心)
    W "after actPlaceD: D=" & Cnt(s, "__pthlp_PD__")
    Dim sh As Shape
    For Each sh In s.Shapes
        If Left$(sh.Name, 8) = "__pthlp_" Then W "   mark " & sh.Name & " at " & Round(sh.Left + sh.Width / 2, 1) & "," & Round(sh.Top + sh.Height / 2, 1)
    Next sh
    Pause 2
    Close #f
End Sub
