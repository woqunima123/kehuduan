Attribute VB_Name = "custom_circles"
Option Explicit

Sub custom_chart_circles(sht As Worksheet, rng As Range, arr, row_start As Integer, col_label As Integer, col_percent As Integer)

    Dim i As Integer
    Dim m_sec As Integer
    Dim shp_circle As Shape, shp_label As Shape, shp_chart As Shape
    Dim wps_office As String
    Dim arr_shapes
    
    ' 先判断一下子是WPS还是office，这俩关于弧线的left和top属性他娘的，不一样
    wps_office = Trim(Application.Caption)
    Do
        wps_office = Replace(wps_office, " ", "")
    Loop While InStr(wps_office, " ") > 0
    
    If InStr(wps_office, "WPS表格") > 0 Then wps_office = "wps" Else wps_office = "excel"
    
    Dim sht_para As Worksheet: Set sht_para = Sheets("bg_paras")
    
    ' 清空现有
    For Each shp_chart In ActiveSheet.Shapes
        If InStr(shp_chart.Name, "chart_") Then shp_chart.Delete
    Next shp_chart
    
    Const INITIAL_SIZE As Integer = 50
    Const INT_MARGIN As Integer = 10
    Const INT_WEIGHT As Integer = 30
    Dim int_count As Integer: int_count = 0
    Dim dbl_top As Double, dbl_left As Double
    Dim int_size As Integer, int_space As Integer
    
    Dim int_r As Integer, int_g As Integer, int_b As Integer
    Dim dark_or_light
    
    int_space = INT_WEIGHT + INT_MARGIN
    
    For i = UBound(arr, 1) To row_start Step -1
    
        dark_or_light = 1000
        ' 改成深色
        Do
            Randomize
            int_r = Int((255 - 0 + 1) * Rnd + 0)
            int_g = Int((255 - 0 + 1) * Rnd + 0)
            int_b = Int((255 - 0 + 1) * Rnd + 0)
            dark_or_light = int_r * 0.299 + int_g * 0.587 + int_b * 0.114
        Loop While dark_or_light > 100
        ' 累计加1
        int_count = int_count + 1
        
        int_size = INITIAL_SIZE + (int_count - 1) * int_space
        If wps_office = "excel" Then
            dbl_left = rng.Left
            dbl_top = rng.Top - (int_count - 1) * int_space - (int_count - 1) * INT_WEIGHT
        Else
            dbl_left = rng.Left - (int_count - 1) * int_space
            dbl_top = rng.Top - (int_count - 1) * int_space
            int_size = int_size * 2
        End If
        
        
        Call copy_circle(sht, 1, CStr(arr(i, col_label)), 0.75, dbl_left, dbl_top, int_size, int_r, int_g, int_b)
        Call copy_circle(sht, CDbl(arr(i, col_percent)), CStr(arr(i, col_label)), 0, dbl_left, dbl_top, int_size, int_r, int_g, int_b)
        
        If wps_office = "excel" Then
            Set shp_label = sht.Shapes.AddShape(msoShapeRectangle, _
                            dbl_left - INT_WEIGHT * 3, _
                            dbl_top - INT_WEIGHT / 2, _
                            INT_WEIGHT * 3, _
                            INT_WEIGHT / 2)
        Else
            Set shp_label = sht.Shapes.AddShape(msoShapeRectangle, _
                            rng.Left - INT_WEIGHT * 3, _
                            dbl_top - INT_WEIGHT / 2, _
                            INT_WEIGHT * 3, _
                            INT_WEIGHT / 2)
        End If
        shp_label.Fill.Visible = msoFalse
        shp_label.Line.Visible = msoFalse
        ' 文本框设置
        With shp_label.TextFrame2
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .TextRange.ParagraphFormat.Alignment = msoAlignRight
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = arr(i, col_label) & " | " & Format(arr(i, col_percent), "0.00%")
        End With
        ' 字体设置
        With shp_label.TextFrame2.TextRange.Font
            .Name = "等线"
            .NameFarEast = "等线"
            .NameAscii = "Consolas"
            .NameOther = "Consolas"
            .Fill.ForeColor.RGB = RGB(int_r, int_g, int_b)
            .Size = 10.5
            .Bold = msoTrue
        End With
        shp_label.Name = "chart_label"
    Next i
    
    ' 将这一堆图形组合起来，方便复制到别处
    arr_shapes = Array("")
    int_count = 0
    For Each shp_chart In ActiveSheet.Shapes
        If InStr(shp_chart.Name, "chart_") Then
            int_count = int_count + 1
            ReDim arr_shapes(1 To int_count)
            arr_shapes(int_count) = shp_chart.Name
        End If
    Next shp_chart
    
    
    
    sht.Range("I9").Select

End Sub

Sub copy_circle(sht As Worksheet, dbl_percent As Double, str_label As String, dbl_transparency As Double, _
                dbl_left As Double, dbl_top As Double, int_size As Integer, _
                int_r As Integer, int_g As Integer, int_b As Integer)

    Dim sht_para As Worksheet: Set sht_para = Sheets("bg_paras")
    Dim shp_circle As Shape
    Dim m_sec As Integer
    
    Const INT_WEIGHT As Integer = 30
    
    Application.ScreenUpdating = False  ' 这里要关掉屏幕刷新，不然复制粘贴那一下子会闪一下，影响观感
    
    Do
        On Error Resume Next
        sht_para.Shapes("chart_circle").Copy
        DoEvents
    Loop While Err.Number <> 0
    
    Do
        On Error Resume Next
        sht.Paste
        DoEvents
    Loop While Err.Number <> 0
    
    Selection.ShapeRange.Name = "chart_circle_" & str_label & CStr(dbl_percent)
    Set shp_circle = sht.Shapes("chart_circle_" & str_label & CStr(dbl_percent))
    sht.Range("I8").Select
    
    shp_circle.Left = dbl_left
    shp_circle.Top = dbl_top
    shp_circle.Height = int_size
    shp_circle.Width = int_size
    
    shp_circle.Fill.Visible = msoFalse
    shp_circle.Line.ForeColor.RGB = RGB(int_r, int_g, int_b)
    shp_circle.Line.Transparency = dbl_transparency
    shp_circle.Line.Weight = INT_WEIGHT
    
    Application.ScreenUpdating = True
    ' 动画部分
    For m_sec = -89 To dbl_percent * 360 - 180
        shp_circle.Adjustments.Item(2) = m_sec
        Call Func_Sleep(2)
    Next m_sec

End Sub

