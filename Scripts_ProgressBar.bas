Attribute VB_Name = "Scripts_ProgressBar"

Sub unload_pbar(sht As Worksheet)
    ' ������н��ȴ���
    Dim shp_pbar As Shape
    
    For Each shp_pbar In sht.Shapes
        If Left(shp_pbar.Name, 5) = "pbar_" Then shp_pbar.Delete
    Next shp_pbar
End Sub
Sub init_pbar(sht As Worksheet, rng As Range, str_title As String, str_text As String, int_current, int_all)
    
    ' *********�Ƴ��ɴ���***********************
    Call unload_pbar(sht)
    sht.Unprotect
    Dim pbar_body_bg As Shape   ' ���屳��
    Dim pbar_title_bg As Shape  ' ���ⱳ��
    Dim pbar_title_text As Shape ' ��������
    Dim pbar_title_red As Shape '��Ȧ
    Dim pbar_title_orange As Shape '��Ȧ
    Dim pbar_title_green As Shape '��Ȧ
    
    Dim pbar_progress_bg As Shape   ' ��������ɫ����
    Dim pbar_progress_fg As Shape   ' ���ȱ���
    
    Dim pbar_body_text As Shape
    Dim pbar_body_info As Shape
    
    Dim dbl_left As Double, dbl_top As Double
    dbl_left = rng.Left
    dbl_top = rng.Top
    
    Const INT_HEIGHT As Integer = 150
    Const INT_WIDTH As Integer = 500
    Const TITLE_HEIGHT As Integer = 25
    ' *********���屳��***********************
    Set pbar_body_bg = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left, _
                    dbl_top, _
                    INT_WIDTH, _
                    INT_HEIGHT)
    pbar_body_bg.Fill.ForeColor.RGB = RGB(255, 255, 255) ' ���ɫǳ��ɫ
    pbar_body_bg.Fill.Transparency = 0.25
    pbar_body_bg.Adjustments.Item(1) = 0.03  ' Բ������
    pbar_body_bg.Line.Visible = msoFalse    ' �ޱ߿�
    pbar_body_bg.Shadow.Type = msoShadow25  ' ��Ӱ����
    pbar_body_bg.Name = "pbar_body_bg"
    
    ' *********������***********************
    Set pbar_title_bg = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left, _
                    dbl_top, _
                    INT_WIDTH, _
                    TITLE_HEIGHT)
    pbar_title_bg.Fill.ForeColor.RGB = RGB(2240, 240, 240) ' ���ɫǳ��ɫ
    pbar_title_bg.Adjustments.Item(1) = 0.2 ' Բ������
    pbar_title_bg.Line.Visible = msoFalse    ' �ޱ߿�
    pbar_title_bg.Name = "pbar_title_bg"
    
    ' --------����������--------------------
    ' ���ֿ���,���ֿ���
    Set pbar_title_text = sht.Shapes.AddShape(msoShapeRoundedRectangle, dbl_left, _
                    dbl_top, _
                    INT_WIDTH, _
                    TITLE_HEIGHT)
    pbar_title_text.Fill.Visible = msoFalse ' �����
    pbar_title_text.TextFrame2.TextRange.Text = str_title
    ' ��������
    pbar_title_text.TextFrame2.TextRange.Font.Name = "����"
    pbar_title_text.TextFrame2.TextRange.Font.NameFarEast = "����"
    pbar_title_text.TextFrame2.TextRange.Font.NameAscii = "Consolas"
    pbar_title_text.TextFrame2.TextRange.Font.NameOther = "Consolas"
    pbar_title_text.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    pbar_title_text.TextFrame2.TextRange.Font.Size = 10.5
    pbar_title_text.TextFrame2.TextRange.Font.Bold = msoTrue
    ' ���뷽ʽ
    pbar_title_text.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
    pbar_title_text.TextFrame2.VerticalAnchor = msoAnchorMiddle
    
    pbar_title_text.Line.Visible = msoFalse
    pbar_title_text.Name = "pbar_title_text"
    
    ' --------����Բ--------------------
    Const SCALE_SIZE As Double = 0.6
    Const LEFT_MARGIN As Integer = 10
    Const ROUND_MARGIN As Integer = 10
    
    Set pbar_title_red = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left + LEFT_MARGIN + TITLE_HEIGHT * SCALE_SIZE * 0 + ROUND_MARGIN * 0, _
                    dbl_top + TITLE_HEIGHT * ((1 - SCALE_SIZE) / 2), _
                    TITLE_HEIGHT * SCALE_SIZE, _
                    TITLE_HEIGHT * SCALE_SIZE)
    pbar_title_red.Line.Visible = msoFalse
    pbar_title_red.Fill.ForeColor.RGB = RGB(255, 79, 79)
    pbar_title_red.Adjustments.Item(1) = 1
    pbar_title_red.Name = "pbar_title_red"
    
    Set pbar_title_orange = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left + LEFT_MARGIN + TITLE_HEIGHT * SCALE_SIZE * 1 + ROUND_MARGIN * 1, _
                    dbl_top + TITLE_HEIGHT * ((1 - SCALE_SIZE) / 2), _
                    TITLE_HEIGHT * SCALE_SIZE, _
                    TITLE_HEIGHT * SCALE_SIZE)
    pbar_title_orange.Line.Visible = msoFalse
    pbar_title_orange.Fill.ForeColor.RGB = RGB(255, 187, 0)
    pbar_title_orange.Adjustments.Item(1) = 1
    pbar_title_orange.Name = "pbar_title_orange"
    
    Set pbar_title_green = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left + LEFT_MARGIN + TITLE_HEIGHT * SCALE_SIZE * 2 + ROUND_MARGIN * 2, _
                    dbl_top + TITLE_HEIGHT * ((1 - SCALE_SIZE) / 2), _
                    TITLE_HEIGHT * SCALE_SIZE, _
                    TITLE_HEIGHT * SCALE_SIZE)
    pbar_title_green.Line.Visible = msoFalse
    pbar_title_green.Fill.ForeColor.RGB = RGB(0, 206, 21)
    pbar_title_green.Adjustments.Item(1) = 1
    pbar_title_green.Name = "pbar_title_green"
    
    
    
    ' *********����������***********************
    Set pbar_progress_bg = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left + LEFT_MARGIN, _
                    dbl_top + TITLE_HEIGHT + LEFT_MARGIN * 3.5, _
                    INT_WIDTH - (LEFT_MARGIN * 2), _
                    TITLE_HEIGHT * 0.5)
    pbar_progress_bg.Line.Visible = msoFalse
    pbar_progress_bg.Fill.ForeColor.RGB = RGB(256, 256, 256)
    pbar_progress_bg.Adjustments.Item(1) = 1
    pbar_progress_bg.Name = "pbar_progress_bg"
    
    ' *********����������***********************
    Set pbar_progress_fg = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left + LEFT_MARGIN, _
                    dbl_top + TITLE_HEIGHT + LEFT_MARGIN * 3.5, _
                    50, _
                    TITLE_HEIGHT * 0.5)
    pbar_progress_fg.Line.Visible = msoFalse
    pbar_progress_fg.Fill.ForeColor.RGB = RGB(255, 79, 79)
    pbar_progress_fg.Adjustments.Item(1) = 1
    pbar_progress_fg.Name = "pbar_progress_fg"
    
    
    ' *********��������***********************
    Set pbar_body_text = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left + LEFT_MARGIN * 2, _
                    pbar_progress_fg.Top + TITLE_HEIGHT * 1.5, _
                    INT_WIDTH - (LEFT_MARGIN), _
                    TITLE_HEIGHT * 0.75)
    pbar_body_text.Line.Visible = msoFalse
    pbar_body_text.Fill.Visible = msoFalse
    pbar_body_text.Adjustments.Item(1) = 0
    pbar_body_text.Name = "pbar_body_text"
    pbar_body_text.TextFrame2.TextRange.Text = str_text
    
    ' ��������
    pbar_body_text.TextFrame2.TextRange.Font.Name = "����"
    pbar_body_text.TextFrame2.TextRange.Font.NameFarEast = "����"
    pbar_body_text.TextFrame2.TextRange.Font.NameAscii = "Consolas"
    pbar_body_text.TextFrame2.TextRange.Font.NameOther = "Consolas"
    pbar_body_text.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(25, 25, 25)
    pbar_body_text.TextFrame2.TextRange.Font.Size = 10.5
    pbar_body_text.TextFrame2.TextRange.Font.Bold = msoTrue
    ' ���뷽ʽ
    pbar_body_text.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    pbar_body_text.TextFrame2.VerticalAnchor = msoAnchorMiddle
    
     ' *********��������***********************
    Set pbar_body_info = sht.Shapes.AddShape(msoShapeRoundedRectangle, _
                    dbl_left + LEFT_MARGIN * 2, _
                    pbar_progress_fg.Top + TITLE_HEIGHT * 2.25, _
                    INT_WIDTH - (LEFT_MARGIN), _
                    TITLE_HEIGHT * 0.75)
    pbar_body_info.Line.Visible = msoFalse
    pbar_body_info.Fill.Visible = msoFalse
    pbar_body_info.Adjustments.Item(1) = 0
    pbar_body_info.Name = "pbar_body_info"
    pbar_body_info.TextFrame2.TextRange.Text = "��  " & int_current & "  ��,��  " & int_all & "  ��"
    
    ' ��������
    pbar_body_info.TextFrame2.TextRange.Font.Name = "����"
    pbar_body_info.TextFrame2.TextRange.Font.NameFarEast = "����"
    pbar_body_info.TextFrame2.TextRange.Font.NameAscii = "Consolas"
    pbar_body_info.TextFrame2.TextRange.Font.NameOther = "Consolas"
    pbar_body_info.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(125, 125, 125)
    pbar_body_info.TextFrame2.TextRange.Font.Size = 10.5
    pbar_body_info.TextFrame2.TextRange.Font.Bold = msoTrue
    ' ���뷽ʽ
    pbar_body_info.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    pbar_body_info.TextFrame2.VerticalAnchor = msoAnchorMiddle
End Sub
Sub set_pbar(sht As Worksheet, str_title As String, str_text As String, int_current, int_all)
    DoEvents
    Dim pbar_title_text As Shape: Set pbar_title_text = sht.Shapes("pbar_title_text") ' ��������
    
    Dim pbar_progress_bg As Shape: Set pbar_progress_bg = sht.Shapes("pbar_progress_bg")    ' ��������ɫ����
    Dim pbar_progress_fg As Shape: Set pbar_progress_fg = sht.Shapes("pbar_progress_fg")    ' ���ȱ���
    
    Dim pbar_body_text As Shape: Set pbar_body_text = sht.Shapes("pbar_body_text")
    Dim pbar_body_info As Shape: Set pbar_body_info = sht.Shapes("pbar_body_info")
    
    ' *********����������***********************
    pbar_progress_fg.ZOrder msoBringToFront
    If int_current / int_all <= 0.33 Then
        pbar_progress_fg.Fill.ForeColor.RGB = RGB(255, 79, 79)
    ElseIf int_current / int_all <= 0.66 Then
        pbar_progress_fg.Fill.ForeColor.RGB = RGB(255, 187, 0)
    Else
        pbar_progress_fg.Fill.ForeColor.RGB = RGB(0, 206, 21)
    End If
    pbar_progress_fg.Width = int_current * pbar_progress_bg.Width / int_all
    
    
    ' *********��������***********************
    pbar_body_text.TextFrame2.TextRange.Text = str_text
    
     ' *********��������***********************
    pbar_body_info.TextFrame2.TextRange.Text = "��  " & int_current & "  ��,��  " & int_all & "  ��"
    
    DoEvents
End Sub


