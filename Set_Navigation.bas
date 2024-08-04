Attribute VB_Name = "Set_Navigation"
Option Explicit

Sub set_pages()

    Dim sht As Worksheet, sht_para As Worksheet
    Dim intR As Integer, intG As Integer, intB As Integer
    Dim rng_title_area As Range, rng_navigation_area As Range
    Dim dic_shts As Object: Set dic_shts = CreateObject("Scripting.Dictionary")
    Dim var_key  As Variant
    
    ' 获取所有工作表列表，这里可以设置跳过哪几个表格
    Set sht_para = Sheets("bg_paras")
    For Each sht In Worksheets
        If sht.Name <> sht_para.Name Then dic_shts(sht.Name) = ""
    Next sht
    
    Dim RGB_1: RGB_1 = RGB(22, 120, 123)
    Dim RGB_2: RGB_2 = RGB(130, 151, 108)
    Dim RGB_3: RGB_3 = RGB(248, 147, 29)
    Dim RGB_4: RGB_4 = RGB(255, 236, 150)
    Dim RGB_5: RGB_5 = RGB(123, 150, 71)
    sht_para.Range("B5") = RGB_1
    ' 设置主题色RGB_1
    
    Application.ScreenUpdating = False
    For Each sht In Worksheets
        If sht.Name <> sht_para.Name Then
            sht.Activate
            sht_para.Range("B4") = sht.Name
            ' *********************全局设置****************************
            
            ' 背景设置
'            sht.SetBackgroundPicture Filename:=ThisWorkbook.Path & "\bg.png"
            sht.Cells.Interior.Color = RGB(242, 242, 242)
            
            ' 去除边框
            Call Set_Borders_None(sht.Cells)
            ' 无合并单元格
            sht.Cells.UnMerge
            
            ' 整体行高为20，列宽为10
            sht.Cells.RowHeight = 20
            sht.Cells.ColumnWidth = 5
            
            ' 单元格对齐方式
            sht.Cells.HorizontalAlignment = xlCenter
            sht.Cells.VerticalAlignment = xlCenter
            
            ' 字体设置
            sht.Cells.Font.Name = "等线"
            sht.Cells.Font.Size = 10
            sht.Cells.Font.Bold = False
            
            ' 标题区域
            Set rng_title_area = sht.Range(sht.Cells(1, 1), sht.Cells(5, 1))
            ' 导航区域
            Set rng_navigation_area = sht.Range(sht.Cells(1, 1), sht.Cells(1, 5))
            
            ' *********************导航栏设置****************************
            
            ' 前五列设置宽度，用于存放导航栏，此处可更改
            rng_navigation_area.EntireColumn.ColumnWidth = 7.5
            ' 两侧窄一点
            sht.Cells(1, 1).EntireColumn.ColumnWidth = 2
            sht.Cells(1, 5).EntireColumn.ColumnWidth = 2
            ' 白色
            rng_navigation_area.EntireColumn.Interior.Color = RGB(240, 240, 240)
            ' 加一条边框线
            With rng_navigation_area.EntireColumn.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = RGB(220, 220, 220)
                .Weight = xlThin
            End With
            
            ' *********************标题栏设置****************************
            
            ' 前五行高度设置为10，用于存放标题，此处可更改
            rng_title_area.EntireRow.RowHeight = 10
            rng_title_area.Font.Size = 30
            rng_title_area.Font.Bold = True
            rng_title_area.HorizontalAlignment = xlLeft
            ' 白色
            rng_title_area.EntireRow.Interior.Color = RGB(240, 240, 240)
            ' 加一条边框线
            With rng_title_area.EntireRow.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = RGB(220, 220, 220)
                .Weight = xlThin
            End With
            
            ' 间隙宽度为2，高度为10
            sht.Cells(1, 6).EntireColumn.ColumnWidth = 2
            sht.Cells(1, 8).EntireColumn.ColumnWidth = 2
            sht.Cells(6, 1).EntireRow.RowHeight = 10
            sht.Cells(7, 1).EntireRow.RowHeight = 10
            
            sht.Cells(1, 7).EntireColumn.ColumnWidth = 0.5
            ' 双色填充
            sht.Range(sht.Cells(7, 7), sht.Cells(Rows.Count, 7)).Select
            With Selection.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 0
                .Gradient.ColorStops.Clear
            End With
            With Selection.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -5.09659108249153E-02
            End With
            With Selection.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.250984221930601
            End With
            
            ' 主题区域白色填充
            sht.Range(sht.Cells(7, 8), sht.Cells(Rows.Count, Columns.Count)).Interior.Color = vbWhite
            
            ' 标题行字体，加粗
            sht.Range(sht.Cells(7, 8), sht.Cells(8, Columns.Count)).Font.Size = 11
            sht.Range(sht.Cells(7, 8), sht.Cells(8, Columns.Count)).Font.Bold = True
            
            ' 网格线关闭
            ActiveWindow.DisplayGridlines = False
            
            ' *********************导航栏标签设置****************************
            
            ' 设置导航栏,这里需要定义几个变量：文本框宽度、右侧指示图标宽度
            Dim nav_label_width As Double: nav_label_width = sht.Range(sht.Cells(1, 1), sht.Cells(1, 5)).Width
            Const NAV_IDENTIFICATION As Double = 5
            Const NAV_LABEL_HEIGHT As Double = 40
            Dim shp As Shape, int_count As Integer, shp_mark As Shape
            ' 清除已有导航标签
            int_count = 0
            For Each shp In sht.Shapes
                If Left(shp.Name, 3) = "nav" Then shp.Delete
            Next shp
            
            ' 更新目录标签
            Set shp = sht.Shapes.AddShape(msoShapeRectangle, _
                                0, _
                                sht.Range("A2:E4").Top, _
                                sht.Range("A2:E4").Width, _
                                sht.Range("A2:E4").Height)
            shp.TextFrame2.TextRange.Text = "更新目录"
            shp.Line.Visible = msoFalse
            shp.Name = "nav_更新目录"
            shp.TextFrame2.TextRange.Font.NameFarEast = "Times New Roman"
            shp.TextFrame2.TextRange.Font.NameOther = "等线"
            shp.TextFrame2.TextRange.Font.Name = "等线"
            shp.TextFrame2.TextRange.Font.Bold = msoTrue
            shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
            shp.OnAction = "set_pages"
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB_1
    
            For Each var_key In dic_shts.keys
                Set shp = sht.Shapes.AddShape(msoShapeRectangle, _
                                0, _
                                NAV_LABEL_HEIGHT * int_count + sht.Range("A1:A7").Height, _
                                nav_label_width, _
                                NAV_LABEL_HEIGHT)
                shp.TextFrame2.TextRange.Text = var_key
                shp.Line.Visible = msoFalse
                shp.Name = "nav_" & var_key
                shp.TextFrame2.TextRange.Font.NameFarEast = "Times New Roman"
                shp.TextFrame2.TextRange.Font.NameOther = "等线"
                shp.TextFrame2.TextRange.Font.Name = "等线"
                shp.TextFrame2.TextRange.Font.Bold = msoTrue
                shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
                shp.OnAction = "turn_to_sht"
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0) ' 先白色填充，然后单独根据是否是本工作表再上主题色
                If sht.Name <> var_key Then
                    shp.Fill.ForeColor.RGB = RGB(240, 240, 240) ' 灰色，这个不变
                Else
                    shp.Fill.Visible = msoTrue
                    ' 插入右侧竖条标志
                    Set shp_mark = sht.Shapes.AddShape(msoShapeRectangle, _
                                shp.Width - NAV_IDENTIFICATION, _
                                shp.Top, _
                                NAV_IDENTIFICATION, _
                                shp.Height)
                    shp_mark.Name = "nav_mark"
                    shp_mark.Line.Visible = msoFalse
                End If
                int_count = int_count + 1
            Next var_key
            
            ' 主题色按钮
            ' 需要计算留出来的空间有多少，然后按照-+-+-+-+-+-的布局方式
            Dim dic_colors As Object: Set dic_colors = CreateObject("Scripting.Dictionary")
            Dim i As Integer
            Dim total_width  As Double: total_width = sht.Range("A1:E1").Width  ' 总宽度
            Const DBL_SPACING As Integer = 10  ' 按钮间距
            Dim dbl_diameter As Double
            
            ' 需要计算每个圆的直径
            dbl_diameter = (total_width - DBL_SPACING * 6) / 5
            
            dic_colors(1) = RGB_1
            dic_colors(2) = RGB_2
            dic_colors(3) = RGB_3
            dic_colors(4) = RGB_4
            dic_colors(5) = RGB_5
            
            For i = 1 To 5
                Set shp = sht.Shapes.AddShape(msoShapeOval, _
                                    dbl_diameter * (i - 1) + DBL_SPACING * i, _
                                    NAV_LABEL_HEIGHT * int_count + sht.Range("A1:A7").Height + 40, _
                                    dbl_diameter, _
                                    dbl_diameter)
                shp.Fill.ForeColor.RGB = dic_colors(i)
                shp.Line.Visible = msoFalse
                shp.Name = "nav_colors_" & i
                shp.OnAction = "set_paras"
            Next i
        End If
    Next sht
    
    For Each sht In Worksheets
        sht.Activate
        sht_para.Range("B4") = sht.Name
        Call Set_Theme
        ActiveWindow.DisplayGridlines = False   ' 网格线不要
        ActiveWindow.DisplayHeadings = False    ' 标题不要
    Next sht
    
    ActiveWindow.DisplayWorkbookTabs = True
    Sheets(1).Activate
    Application.ScreenUpdating = True
    
    MsgBox "Done！"

End Sub

Sub set_paras()

    Dim sht As Worksheet: Set sht = ActiveSheet
    Dim shp As Shape: Set shp = sht.Shapes(Application.Caller)
    Dim sht_para As Worksheet: Set sht_para = Sheets("bg_paras")
    
    sht_para.Range("B4") = sht.Name
    sht_para.Range("B5") = shp.Fill.ForeColor.RGB
    
    Call Set_Theme

End Sub

Sub Set_Theme()

    Dim shp As Shape
    Dim r As Integer, g As Integer, b As Integer
    Dim light_or_dark As Double
    Dim sht_para As Worksheet: Set sht_para = Sheets("bg_paras")
    Dim sht As Worksheet: Set sht = Sheets(sht_para.Range("B4").Text)
    Dim rgb_theme: rgb_theme = sht_para.Range("B5")
    
    If sht.Name = sht_para.Name Then Exit Sub
    
    Dim i As Integer
    
    r = CLng(rgb_theme) Mod 256
    g = CLng(rgb_theme) / 256 Mod 256
    b = CLng(rgb_theme) / 65536 Mod 256
    light_or_dark = r * 0.229 + g * 0.587 + b * 0.114
    
    ' 标题文字部分
    sht.Range(sht.Cells(1, 6), sht.Cells(5, Columns.Count)).Font.Size = 18
    sht.Range(sht.Cells(1, 6), sht.Cells(5, Columns.Count)).Font.Bold = True
    sht.Range(sht.Cells(1, 6), sht.Cells(5, Columns.Count)).Font.Color = rgb_theme
    sht.Range(sht.Cells(1, 6), sht.Cells(5, Columns.Count)).HorizontalAlignment = xlLeft
    sht.Range(sht.Cells(2, 9), sht.Cells(4, 20)).Merge  ' 合并单元格
    
    ' 正文区域格式设置
    ' 正文从I8单元格开始，所以这个单元格不能为空，并且至少要是两个单元格组成一个连续区域
    Dim int_max_column As Integer
    
    If sht.Range("I8") <> "" And sht.Range("I8").CurrentRegion.Cells.Count > 1 Then
        
        int_max_column = sht.Range("I8").CurrentRegion.Columns.Count + 8
    
        Call Set_Borders_None(sht.Range(sht.Range("I8"), sht.Cells(10000, int_max_column)))
        
        sht.Range(sht.Cells(8, 9), sht.Cells(8, int_max_column)).EntireColumn.AutoFit
        For i = 9 To int_max_column
            sht.Cells(8, i).EntireColumn.AutoFit
            sht.Cells(8, i).EntireColumn.ColumnWidth = sht.Cells(8, i).EntireColumn.ColumnWidth + 1
        Next i
        
        ' 标题行主题色
        sht.Range(sht.Cells(8, 9), sht.Cells(8, int_max_column)).Interior.Color = rgb_theme
        
        If light_or_dark < 100 Then
            sht.Range(sht.Cells(8, 9), sht.Cells(8, int_max_column)).Font.Color = RGB(255, 255, 255)
        Else
            sht.Range(sht.Cells(8, 9), sht.Cells(8, int_max_column)).Font.Color = RGB(0, 0, 0)
        End If
        
        ' 边框
        Call Set_Borders_All(sht.Range("I8").CurrentRegion, rgb_theme)
        
    End If
    
    
    For Each shp In sht.Shapes
        If Left(shp.Name, 3) = "nav" And InStr(shp.Name, "color") = 0 Then
            shp.Fill.ForeColor.RGB = RGB(240, 240, 240)
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0) ' 白底黑字
            shp.TextFrame2.TextRange.Font.Bold = msoFalse  ' 字体不加粗
        End If
        ' 当前工作表
        If shp.Name = "nav_" & sht.Name Or shp.Name = "nav_更新目录" Then
            shp.Fill.ForeColor.RGB = rgb_theme  ' 主题色
            shp.Fill.Transparency = 0.5
            If light_or_dark < 100 Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Else
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End If
            shp.TextFrame2.TextRange.Font.Bold = msoTrue ' 字体加粗
        End If
    Next shp
    
    ' 小竖条
    Set shp = sht.Shapes("nav_mark")
    shp.Fill.ForeColor.RGB = rgb_theme
    shp.Fill.Transparency = 0
    
    ' ************************ 有新增的功能，加到下面就行 *************************
    ' 按钮
    For Each shp In sht.Shapes
        If Left(shp.Name, 4) = "btn_" Then
            shp.Fill.ForeColor.RGB = rgb_theme  ' 主题色
            shp.Fill.Transparency = 0
            If light_or_dark < 100 Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Else
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End If
            shp.TextFrame2.TextRange.Font.Bold = msoTrue ' 字体加粗
        End If
    Next shp
    
    sht.Cells(1, 1).Select
    
    ActiveWindow.Zoom = 100 ' 缩放100

End Sub

Sub turn_to_sht()

    Dim shp As Shape
    Dim sht As Worksheet
    
    Set shp = ActiveSheet.Shapes(Application.Caller)
    
    Set sht = Sheets(shp.TextFrame2.TextRange.Text)
    sht.Activate
    sht.Cells(1, 1).Select

End Sub
