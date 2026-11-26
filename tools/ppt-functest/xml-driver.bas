Attribute VB_Name = "modXmlDriver"
Option Explicit

' 走产品真实入口生成 直线/抛物线/圆弧 三种路径动画, 另存副本供外部解包核对 animMotion XML

Public Sub FT_Xml()
    On Error GoTo eh
    Dim f As Integer
    f = FreeFile
    Open Environ$("TEMP") & "\abp_ft_results.txt" For Output As #f
    Print #f, "xml gen start"
    Close #f
    Dim pres As Presentation: Set pres = ActivePresentation
    Do While pres.Slides.Count > 1
        pres.Slides(pres.Slides.Count).Delete
    Loop
    If pres.Slides.Count = 0 Then pres.Slides.Add 1, ppLayoutBlank
    Dim s As Slide: Set s = pres.Slides(1)
    Do While s.Shapes.Count > 0
        s.Shapes(1).Delete
    Loop
    Do While s.TimeLine.MainSequence.Count > 0
        s.TimeLine.MainSequence(1).Delete
    Loop
    Dim sh As Shape: Set sh = s.Shapes.AddShape(msoShapeRectangle, 100!, 100!, 120!, 80!)
    sh.Name = "FT_mover"
    Dim tgt As Shape: Set tgt = s.Shapes.AddShape(msoShapeOval, 500!, 300!, 100!, 100!)
    tgt.Name = "FT_target"

    ' ---- 直线: B 拖到目标中心 ----
    sh.Select
    Application.Run "actStart"
    Dim bm As Shape
    Set bm = FindMarkPub(s, "__pthlp_B__")
    If Not bm Is Nothing Then
        bm.Left = 550! - bm.Width / 2!: bm.Top = 350! - bm.Height / 2!
    End If
    Application.Run "actGenerate"
    f = FreeFile
    Open Environ$("TEMP") & "\abp_ft_results.txt" For Append As #f
    Print #f, "line effects=" & s.TimeLine.MainSequence.Count
    Close #f

    ' ---- 抛物线 (P 顶点 300,60; D 终点在目标中心) ----
    Application.Run "SetForm", "抛物线"
    sh.Select
    Application.Run "actStart"
    tgt.Select
    Application.Run "actPlaceD"
    s.Shapes.AddShape msoShapeCross, 300! - 3.5!, 60! - 3.5!, 7!, 7!
    Dim pm As Shape
    Set pm = s.Shapes(s.Shapes.Count)
    pm.Name = "__pthlp_PM__"
    Application.Run "actGenerate"
    f = FreeFile
    Open Environ$("TEMP") & "\abp_ft_results.txt" For Append As #f
    Print #f, "para effects=" & s.TimeLine.MainSequence.Count
    Close #f

    ' ---- 圆弧 (M 弧点 420,120; D 重新放到目标中心) ----
    Application.Run "SetForm", "圆弧"
    sh.Select
    Application.Run "actStart"
    tgt.Select
    Application.Run "actPlaceD"
    s.Shapes.AddShape msoShapeCross, 420! - 3.5!, 120! - 3.5!, 7!, 7!
    Set pm = s.Shapes(s.Shapes.Count)
    pm.Name = "__pthlp_PM__"
    Application.Run "actGenerate"
    f = FreeFile
    Open Environ$("TEMP") & "\abp_ft_results.txt" For Append As #f
    Print #f, "arc effects=" & s.TimeLine.MainSequence.Count
    Close #f

    pres.SaveCopyAs Environ$("TEMP") & "\abp_xml.pptx", ppSaveAsOpenXMLPresentation
    f = FreeFile
    Open Environ$("TEMP") & "\abp_ft_results.txt" For Append As #f
    Print #f, "SAVED"
    Close #f
    Exit Sub
eh:
    f = FreeFile
    Open Environ$("TEMP") & "\abp_ft_results.txt" For Append As #f
    Print #f, "ERR " & Err.Number & " " & Err.Description
    Close #f
End Sub

' 找加载项的标记
Private Function FindMarkPub(s As Slide, nm As String) As Shape
    Dim i As Long
    For i = 1 To s.Shapes.Count
        If s.Shapes(i).Name = nm Then Set FindMarkPub = s.Shapes(i): Exit Function
    Next i
End Function
