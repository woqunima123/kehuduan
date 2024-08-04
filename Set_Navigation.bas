Attribute VB_Name = "Set_Navigation"
Option Explicit

Sub set_pages()

    Dim sht As Worksheet, sht_para As Worksheet
    Dim intR As Integer, intG As Integer, intB As Integer
    Dim rng_title_area As Range, rng_navigation_area As Range
    Dim dic_shts As Object: Set dic_shts = CreateObject("Scripting.Dictionary")
    Dim var_key  As Variant
    
    ' ��ȡ���й������б�����������������ļ������
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
    ' ��������ɫRGB_1
    
    Application.ScreenUpdating = False
    For Each sht In Worksheets
        If sht.Name <> sht_para.Name Then
            sht.Activate
            sht_para.Range("B4") = sht.Name
            ' *********************ȫ������****************************
            
            ' ��������
'            sht.SetBackgroundPicture Filename:=ThisWorkbook.Path & "\bg.png"
            sht.Cells.Interior.Color = RGB(242, 242, 242)
            
            ' ȥ���߿�
            Call Set_Borders_None(sht.Cells)
            ' �޺ϲ���Ԫ��
            sht.Cells.UnMerge
            
            ' �����и�Ϊ20���п�Ϊ10
            sht.Cells.RowHeight = 20
            sht.Cells.ColumnWidth = 5
            
            ' ��Ԫ����뷽ʽ
            sht.Cells.HorizontalAlignment = xlCenter
            sht.Cells.VerticalAlignment = xlCenter
            
            ' ��������
            sht.Cells.Font.Name = "����"
            sht.Cells.Font.Size = 10
            sht.Cells.Font.Bold = False
            
            ' ��������
            Set rng_title_area = sht.Range(sht.Cells(1, 1), sht.Cells(5, 1))
            ' ��������
            Set rng_navigation_area = sht.Range(sht.Cells(1, 1), sht.Cells(1, 5))
            
            ' *********************����������****************************
            
            ' ǰ�������ÿ�ȣ����ڴ�ŵ��������˴��ɸ���
            rng_navigation_area.EntireColumn.ColumnWidth = 7.5
            ' ����խһ��
            sht.Cells(1, 1).EntireColumn.ColumnWidth = 2
            sht.Cells(1, 5).EntireColumn.ColumnWidth = 2
            ' ��ɫ
            rng_navigation_area.EntireColumn.Interior.Color = RGB(240, 240, 240)
            ' ��һ���߿���
            With rng_navigation_area.EntireColumn.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = RGB(220, 220, 220)
                .Weight = xlThin
            End With
            
            ' *********************����������****************************
            
            ' ǰ���и߶�����Ϊ10�����ڴ�ű��⣬�˴��ɸ���
            rng_title_area.EntireRow.RowHeight = 10
            rng_title_area.Font.Size = 30
            rng_title_area.Font.Bold = True
            rng_title_area.HorizontalAlignment = xlLeft
            ' ��ɫ
            rng_title_area.EntireRow.Interior.Color = RGB(240, 240, 240)
            ' ��һ���߿���
            With rng_title_area.EntireRow.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = RGB(220, 220, 220)
                .Weight = xlThin
            End With
            
            ' ��϶���Ϊ2���߶�Ϊ10
            sht.Cells(1, 6).EntireColumn.ColumnWidth = 2
            sht.Cells(1, 8).EntireColumn.ColumnWidth = 2
            sht.Cells(6, 1).EntireRow.RowHeight = 10
            sht.Cells(7, 1).EntireRow.RowHeight = 10
            
            sht.Cells(1, 7).EntireColumn.ColumnWidth = 0.5
            ' ˫ɫ���
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
            
            ' ���������ɫ���
            sht.Range(sht.Cells(7, 8), sht.Cells(Rows.Count, Columns.Count)).Interior.Color = vbWhite
            
            ' ���������壬�Ӵ�
            sht.Range(sht.Cells(7, 8), sht.Cells(8, Columns.Count)).Font.Size = 11
            sht.Range(sht.Cells(7, 8), sht.Cells(8, Columns.Count)).Font.Bold = True
            
            ' �����߹ر�
            ActiveWindow.DisplayGridlines = False
            
            ' *********************��������ǩ����****************************
            
            ' ���õ�����,������Ҫ���弸���������ı����ȡ��Ҳ�ָʾͼ����
            Dim nav_label_width As Double: nav_label_width = sht.Range(sht.Cells(1, 1), sht.Cells(1, 5)).Width
            Const NAV_IDENTIFICATION As Double = 5
            Const NAV_LABEL_HEIGHT As Double = 40
            Dim shp As Shape, int_count As Integer, shp_mark As Shape
            ' ������е�����ǩ
            int_count = 0
            For Each shp In sht.Shapes
                If Left(shp.Name, 3) = "nav" Then shp.Delete
            Next shp
            
            ' ����Ŀ¼��ǩ
            Set shp = sht.Shapes.AddShape(msoShapeRectangle, _
                                0, _
                                sht.Range("A2:E4").Top, _
                                sht.Range("A2:E4").Width, _
                                sht.Range("A2:E4").Height)
            shp.TextFrame2.TextRange.Text = "����Ŀ¼"
            shp.Line.Visible = msoFalse
            shp.Name = "nav_����Ŀ¼"
            shp.TextFrame2.TextRange.Font.NameFarEast = "Times New Roman"
            shp.TextFrame2.TextRange.Font.NameOther = "����"
            shp.TextFrame2.TextRange.Font.Name = "����"
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
                shp.TextFrame2.TextRange.Font.NameOther = "����"
                shp.TextFrame2.TextRange.Font.Name = "����"
                shp.TextFrame2.TextRange.Font.Bold = msoTrue
                shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
                shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
                shp.OnAction = "turn_to_sht"
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0) ' �Ȱ�ɫ��䣬Ȼ�󵥶������Ƿ��Ǳ���������������ɫ
                If sht.Name <> var_key Then
                    shp.Fill.ForeColor.RGB = RGB(240, 240, 240) ' ��ɫ���������
                Else
                    shp.Fill.Visible = msoTrue
                    ' �����Ҳ�������־
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
            
            ' ����ɫ��ť
            ' ��Ҫ�����������Ŀռ��ж��٣�Ȼ����-+-+-+-+-+-�Ĳ��ַ�ʽ
            Dim dic_colors As Object: Set dic_colors = CreateObject("Scripting.Dictionary")
            Dim i As Integer
            Dim total_width  As Double: total_width = sht.Range("A1:E1").Width  ' �ܿ��
            Const DBL_SPACING As Integer = 10  ' ��ť���
            Dim dbl_diameter As Double
            
            ' ��Ҫ����ÿ��Բ��ֱ��
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
        ActiveWindow.DisplayGridlines = False   ' �����߲�Ҫ
        ActiveWindow.DisplayHeadings = False    ' ���ⲻҪ
    Next sht
    
    ActiveWindow.DisplayWorkbookTabs = True
    Sheets(1).Activate
    Application.ScreenUpdating = True
    
    MsgBox "Done��"

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
    
    ' �������ֲ���
    sht.Range(sht.Cells(1, 6), sht.Cells(5, Columns.Count)).Font.Size = 18
    sht.Range(sht.Cells(1, 6), sht.Cells(5, Columns.Count)).Font.Bold = True
    sht.Range(sht.Cells(1, 6), sht.Cells(5, Columns.Count)).Font.Color = rgb_theme
    sht.Range(sht.Cells(1, 6), sht.Cells(5, Columns.Count)).HorizontalAlignment = xlLeft
    sht.Range(sht.Cells(2, 9), sht.Cells(4, 20)).Merge  ' �ϲ���Ԫ��
    
    ' ���������ʽ����
    ' ���Ĵ�I8��Ԫ��ʼ�����������Ԫ����Ϊ�գ���������Ҫ��������Ԫ�����һ����������
    Dim int_max_column As Integer
    
    If sht.Range("I8") <> "" And sht.Range("I8").CurrentRegion.Cells.Count > 1 Then
        
        int_max_column = sht.Range("I8").CurrentRegion.Columns.Count + 8
    
        Call Set_Borders_None(sht.Range(sht.Range("I8"), sht.Cells(10000, int_max_column)))
        
        sht.Range(sht.Cells(8, 9), sht.Cells(8, int_max_column)).EntireColumn.AutoFit
        For i = 9 To int_max_column
            sht.Cells(8, i).EntireColumn.AutoFit
            sht.Cells(8, i).EntireColumn.ColumnWidth = sht.Cells(8, i).EntireColumn.ColumnWidth + 1
        Next i
        
        ' ����������ɫ
        sht.Range(sht.Cells(8, 9), sht.Cells(8, int_max_column)).Interior.Color = rgb_theme
        
        If light_or_dark < 100 Then
            sht.Range(sht.Cells(8, 9), sht.Cells(8, int_max_column)).Font.Color = RGB(255, 255, 255)
        Else
            sht.Range(sht.Cells(8, 9), sht.Cells(8, int_max_column)).Font.Color = RGB(0, 0, 0)
        End If
        
        ' �߿�
        Call Set_Borders_All(sht.Range("I8").CurrentRegion, rgb_theme)
        
    End If
    
    
    For Each shp In sht.Shapes
        If Left(shp.Name, 3) = "nav" And InStr(shp.Name, "color") = 0 Then
            shp.Fill.ForeColor.RGB = RGB(240, 240, 240)
            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0) ' �׵׺���
            shp.TextFrame2.TextRange.Font.Bold = msoFalse  ' ���岻�Ӵ�
        End If
        ' ��ǰ������
        If shp.Name = "nav_" & sht.Name Or shp.Name = "nav_����Ŀ¼" Then
            shp.Fill.ForeColor.RGB = rgb_theme  ' ����ɫ
            shp.Fill.Transparency = 0.5
            If light_or_dark < 100 Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Else
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End If
            shp.TextFrame2.TextRange.Font.Bold = msoTrue ' ����Ӵ�
        End If
    Next shp
    
    ' С����
    Set shp = sht.Shapes("nav_mark")
    shp.Fill.ForeColor.RGB = rgb_theme
    shp.Fill.Transparency = 0
    
    ' ************************ �������Ĺ��ܣ��ӵ�������� *************************
    ' ��ť
    For Each shp In sht.Shapes
        If Left(shp.Name, 4) = "btn_" Then
            shp.Fill.ForeColor.RGB = rgb_theme  ' ����ɫ
            shp.Fill.Transparency = 0
            If light_or_dark < 100 Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Else
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End If
            shp.TextFrame2.TextRange.Font.Bold = msoTrue ' ����Ӵ�
        End If
    Next shp
    
    sht.Cells(1, 1).Select
    
    ActiveWindow.Zoom = 100 ' ����100

End Sub

Sub turn_to_sht()

    Dim shp As Shape
    Dim sht As Worksheet
    
    Set shp = ActiveSheet.Shapes(Application.Caller)
    
    Set sht = Sheets(shp.TextFrame2.TextRange.Text)
    sht.Activate
    sht.Cells(1, 1).Select

End Sub
