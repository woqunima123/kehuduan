Attribute VB_Name = "Scripts_Messagebox"
Option Explicit

Sub Load_Messagebox(sht As Worksheet, rng As Range, str_type As String, str_text As String)
    
    Application.ScreenUpdating = False
    
    ' 直接把sht写入参数表
    Sheets("bg_paras").Range("B2") = sht.Name
    Call Unload_Messagebox_Yes
    Sheets("bg_paras").Range("B3") = ""
    sht.Unprotect
    Dim msgbox_body_bg As Shape
    Dim msgbox_body_logo As Shape
    Dim msgbox_body_text As Shape
    Dim msgbox_button_yes As Shape
    Dim msgbox_button_no As Shape
    
    Const INT_HEIGHT As Integer = 175   ' 窗体总高度
    Const INT_WIDTH As Integer = 450 ' 窗体总宽度
    Const MARGIN_VER As Integer = 10
    Const MARGIN_HOR As Integer = 15
    
    Dim dbl_left As Double: dbl_left = rng.Left
    Dim dbl_top As Double: dbl_top = rng.Top
    
    
    Set msgbox_body_bg = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left, _
                    dbl_top, _
                    INT_WIDTH, _
                    INT_HEIGHT)
    msgbox_body_bg.Fill.ForeColor.RGB = RGB(245, 245, 245) ' 填充色浅灰色
    msgbox_body_bg.Fill.Transparency = 0.1
    msgbox_body_bg.Adjustments.Item(1) = 0.05 ' 圆角设置
    msgbox_body_bg.Line.Visible = msoFalse    ' 无边框
    msgbox_body_bg.Name = "msgbox_body_bg"
    
    ' 阴影设置
    With msgbox_body_bg.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 5
        .OffsetX = 1
        .OffsetY = 1
        .RotateWithShape = msoFalse
        .ForeColor.RGB = RGB(16, 120, 253)
        .Transparency = 0.85
        .Size = 100.5
    End With
    
    Const INT_LOGO_SIZE As Integer = 75
    
    Do
        On Error Resume Next
        Sheets("bg_paras").Shapes(str_type).Copy
        DoEvents
    Loop While Err.Number <> 0
    
    Do
        On Error Resume Next
        sht.Paste
        DoEvents
    Loop While Err.Number <> 0
    
    Selection.ShapeRange.Name = "msgbox_body_logo"
    Set msgbox_body_logo = sht.Shapes("msgbox_body_logo")
    msgbox_body_logo.Left = dbl_left + msgbox_body_bg.Width / 2 - INT_LOGO_SIZE / 2
    msgbox_body_logo.Top = dbl_top + MARGIN_VER
    msgbox_body_logo.Width = INT_LOGO_SIZE
    msgbox_body_logo.Height = INT_LOGO_SIZE
    
    Set msgbox_body_text = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left + MARGIN_HOR * 2, _
                    msgbox_body_logo.Top + msgbox_body_logo.Height + MARGIN_VER * 2, _
                    INT_WIDTH - (MARGIN_HOR * 4), _
                    40)
    msgbox_body_text.Name = "msgbox_body_text"
    msgbox_body_text.Fill.Visible = msoFalse
    msgbox_body_text.Line.Visible = msoFalse
    ' 字体设置
    With msgbox_body_text.TextFrame2.TextRange.Font
        .Name = "等线"
        .NameFarEast = "等线"
        .NameAscii = "Consolas"
        .NameOther = "Consolas"
        .Fill.ForeColor.RGB = RGB(50, 50, 50)
        .Size = 11
        .Bold = msoTrue
    End With
    ' 对齐方式
    msgbox_body_text.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    msgbox_body_text.TextFrame2.VerticalAnchor = msoAnchorTop
    msgbox_body_text.TextFrame2.TextRange.Text = str_text
    
    Const BTN_HEIGHT As Integer = 20
    Const BTN_WIDTH As Integer = 60
    
    ' 右下角按钮 YES
    Set msgbox_button_yes = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left + msgbox_body_bg.Width - MARGIN_HOR - BTN_WIDTH, _
                    msgbox_body_bg.Top + msgbox_body_bg.Height - MARGIN_VER * 1 - BTN_HEIGHT, _
                    BTN_WIDTH, _
                    BTN_HEIGHT)
    msgbox_button_yes.Name = "msgbox_button_yes"
    If str_type = "information" Then
        msgbox_button_yes.Fill.ForeColor.RGB = RGB(0, 206, 21)
    ElseIf str_type = "warning" Then
        msgbox_button_yes.Fill.ForeColor.RGB = RGB(255, 79, 79)
    ElseIf str_type = "question" Then
        msgbox_button_yes.Fill.ForeColor.RGB = RGB(255, 187, 0)
    End If
    msgbox_button_yes.Adjustments.Item(1) = 0.25
    msgbox_button_yes.Line.Visible = msoFalse
    ' 字体设置
    With msgbox_button_yes.TextFrame2.TextRange.Font
        .Name = "等线"
        .NameFarEast = "等线"
        .NameAscii = "Consolas"
        .NameOther = "Consolas"
        .Fill.ForeColor.RGB = RGB(50, 50, 50)
        .Size = 11
        .Bold = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    ' 对齐方式
    msgbox_button_yes.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    msgbox_button_yes.TextFrame2.VerticalAnchor = msoAnchorMiddle
    msgbox_button_yes.TextFrame2.TextRange.Text = "Yes"
    
    msgbox_button_yes.OnAction = "Unload_Messagebox_Yes"
    ' 右下角按钮 No
    If str_type = "question" Then
        Set msgbox_button_no = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                        dbl_left + msgbox_body_bg.Width - MARGIN_HOR * 2 - BTN_WIDTH * 2, _
                        msgbox_body_bg.Top + msgbox_body_bg.Height - MARGIN_VER * 1 - BTN_HEIGHT, _
                        BTN_WIDTH, _
                        BTN_HEIGHT)
        msgbox_button_no.Name = "msgbox_button_yes"
        msgbox_button_no.Fill.ForeColor.RGB = RGB(7, 87, 224)
        msgbox_button_no.Adjustments.Item(1) = 0.25
        msgbox_button_no.Line.Visible = msoFalse
        ' 字体设置
        With msgbox_button_no.TextFrame2.TextRange.Font
            .Name = "等线"
            .NameFarEast = "等线"
            .NameAscii = "Consolas"
            .NameOther = "Consolas"
            .Fill.ForeColor.RGB = RGB(50, 50, 50)
            .Size = 11
            .Bold = msoTrue
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
        ' 对齐方式
        msgbox_button_no.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        msgbox_button_no.TextFrame2.VerticalAnchor = msoAnchorMiddle
        msgbox_button_no.TextFrame2.TextRange.Text = "No"
        
        msgbox_button_no.OnAction = "Unload_Messagebox_No"
        
    End If
    
    sht.Range("A1").Select
    
    sht.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Application.ScreenUpdating = True
    
End Sub
Sub Unload_Messagebox_Yes()
    Dim shp As Shape
    Dim sht As Worksheet: Set sht = Sheets(Sheets("bg_paras").Range("B2").Text)
    
    sht.Unprotect
    
    For Each shp In sht.Shapes
        If Left(shp.Name, 7) = "msgbox_" Then shp.Delete
    Next shp
    
    Sheets("bg_paras").Range("B3") = "yes"
    
End Sub
Sub Unload_Messagebox_No()
    Dim shp As Shape
    Dim sht As Worksheet: Set sht = Sheets(Sheets("bg_paras").Range("B2").Text)
    
    sht.Unprotect
    
    For Each shp In sht.Shapes
        If Left(shp.Name, 7) = "msgbox_" Then shp.Delete
    Next shp
    
    Sheets("bg_paras").Range("B3") = "no"
    
End Sub


