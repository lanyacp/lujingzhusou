Attribute VB_Name = "modPathSimple"
Option Explicit

' ===================================================================
'  路径助手 · 简易宏版（PowerPoint）
'  不依赖加载项/窗体/Ribbon，杜绝编译错误。
'
'  用法：
'   1) 打开你的 .pptm/.pptx，Alt+F11 → 菜单"文件→导入文件"选本 .bas
'      （或 插入→模块，把本文件内容粘贴进去）
'   2) 回到幻灯片：先点选"要移动的图形"，再按住 Ctrl 点选"目标图形"
'      （按这个顺序选两个图形）
'   3) Alt+F8 → 运行 "精确路径_中心到中心"
'   → 生成从第一个图形中心到第二个图形中心的精确直线动画
'
'  其他：
'   - "精确路径_到坐标" ：只选1个图形，弹框输入目标坐标(磅)
'   - "删除本图形动画" ：删掉选中图形上已有的动画，方便重来
' ===================================================================

Public Sub 精确路径_中心到中心()
    On Error GoTo eh
    ' 空选择时读 ShapeRange 会抛无法容错的 COM 错误, 必须先判 Type
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "请按顺序选中【两个】图形：先选要移动的，再 Ctrl+选目标。", vbExclamation, "路径助手"
        Exit Sub
    End If
    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange
    If sr.Count <> 2 Then
        MsgBox "请按顺序选中【两个】图形：先选要移动的，再 Ctrl+选目标。", vbExclamation, "路径助手"
        Exit Sub
    End If
    Dim mover As Shape, target As Shape
    Set mover = sr(1)
    Set target = sr(2)
    ' 选区按层叠顺序返回, 与点击顺序可能相反: 让用户确认方向(否=对调, 取消=不做)
    Dim r As VbMsgBoxResult
    r = MsgBox("将生成路径:  [" & mover.Name & "] 中心  →  [" & target.Name & "] 中心" & vbCrLf & vbCrLf & _
               "方向对吗?  【是】生成   【否】对调方向   【取消】不做", _
               vbYesNoCancel + vbQuestion, "路径助手")
    If r = vbCancel Then Exit Sub
    If r = vbNo Then
        Dim tmp As Shape: Set tmp = mover: Set mover = target: Set target = tmp
    End If
    Dim sw As Single, shh As Single
    sw = ActivePresentation.PageSetup.SlideWidth
    shh = ActivePresentation.PageSetup.SlideHeight
    Dim ax As Single, ay As Single, bx As Single, by As Single
    ax = mover.Left + mover.Width / 2!: ay = mover.Top + mover.Height / 2!
    bx = target.Left + target.Width / 2!: by = target.Top + target.Height / 2!
    AddLinePath mover, (bx - ax) / sw, (by - ay) / shh
    MsgBox "已生成精确路径：" & mover.Name & " 中心 → " & target.Name & " 中心", vbInformation, "路径助手"
    Exit Sub
eh: MsgBox "失败[" & Err.Number & "]: " & Err.Description, vbExclamation, "路径助手"
End Sub

Public Sub 精确路径_到坐标()
    On Error GoTo eh
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "请只选中【一个】要移动的图形。", vbExclamation, "路径助手"
        Exit Sub
    End If
    Dim sr As ShapeRange
    Set sr = ActiveWindow.Selection.ShapeRange
    If sr.Count <> 1 Then
        MsgBox "请只选中【一个】要移动的图形。", vbExclamation, "路径助手"
        Exit Sub
    End If
    Dim mover As Shape: Set mover = sr(1)
    Dim tx As String, ty As String, xx As Single, yy As Single
    tx = InputBox("目标点 X 坐标（磅，0~" & ActivePresentation.PageSetup.SlideWidth & "）", "路径助手", "600")
    ty = InputBox("目标点 Y 坐标（磅，0~" & ActivePresentation.PageSetup.SlideHeight & "）", "路径助手", "300")
    If Not IsNumeric(tx) Or Not IsNumeric(ty) Then Exit Sub   ' 取消/无效输入: 安静退出
    xx = CSng(tx): yy = CSng(ty)
    Dim sw As Single, shh As Single
    sw = ActivePresentation.PageSetup.SlideWidth: shh = ActivePresentation.PageSetup.SlideHeight
    Dim ax As Single, ay As Single
    ax = mover.Left + mover.Width / 2!: ay = mover.Top + mover.Height / 2!
    AddLinePath mover, (xx - ax) / sw, (yy - ay) / shh
    MsgBox "已生成到坐标(" & xx & "," & yy & ")的精确路径", vbInformation, "路径助手"
    Exit Sub
eh: MsgBox "失败[" & Err.Number & "]: " & Err.Description, vbExclamation, "路径助手"
End Sub

Public Sub 删除本图形动画()
    On Error Resume Next
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    Dim sr As ShapeRange: Set sr = ActiveWindow.Selection.ShapeRange
    Dim i As Long
    For i = sr.Count To 1 Step -1
        Dim s As Slide: Set s = sr(i).Parent
        Dim seq As Sequence: Set seq = s.TimeLine.MainSequence
        Dim j As Long
        For j = seq.Count To 1 Step -1
            If seq(j).Shape.Name = sr(i).Name Then seq(j).Delete
        Next j
    Next i
    MsgBox "已删除选中图形的动画", vbInformation, "路径助手"
End Sub

' 核心：生成一条精确直线位移动画（起点=图形中心）
Private Sub AddLinePath(sh As Shape, dx As Single, dy As Single)
    Dim s As Slide: Set s = sh.Parent
    Dim eff As Effect
    Set eff = s.TimeLine.MainSequence.AddEffect(sh, msoAnimEffectCustom, , msoAnimTriggerOnPageClick)
    Dim beh As AnimationBehavior
    Set beh = eff.Behaviors.Add(msoAnimTypeMotion)
    beh.MotionEffect.FromX = 0: beh.MotionEffect.FromY = 0
    beh.MotionEffect.ToX = dx: beh.MotionEffect.ToY = dy
    eff.Timing.Duration = 2!
End Sub
