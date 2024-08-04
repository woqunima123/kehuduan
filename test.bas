Attribute VB_Name = "test"
Option Explicit

Sub test_progressbar()

    ActiveSheet.Unprotect
    Call Load_Messagebox(ActiveSheet, Range("I9"), "question", "���г���ǰ������������������ԣ�")
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Do While Sheets("bg_paras").Range("B3") = ""
        DoEvents
    Loop
    If Sheets("bg_paras").Range("B3") = "no" Then
        ActiveSheet.Unprotect
        Call Load_Messagebox(ActiveSheet, Range("I9"), "information", "����ֹ�˳��������")
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        Exit Sub
    End If
    
    Call init_pbar(ActiveSheet, Range("I9"), "���Դ��� | for test", "׼�������� Ready... ...", 0, 100)
    
    Call Func_Sleep(1000)
    
    Dim i As Integer
    
    For i = 1 To 100
        Call set_pbar(ActiveSheet, " you shall not pass", "�������������С����� ������", i, 100)
        Call Func_Sleep(25)
    Next i
    
    Call unload_pbar(ActiveSheet)
    
    Call Load_Messagebox(ActiveSheet, Range("I9"), "information", "�㶨������������ݣ�")
    
End Sub

Sub test_circles()

    Dim arr: arr = Range("I8").CurrentRegion
    
    Calculate
    
    Call custom_chart_circles(ActiveSheet, ActiveSheet.Range("Q15"), arr, 2, 1, 2)

End Sub

Sub get_test()

    Dim word_app As Word.Application
    Set word_app = New Word.Application
    word_app.Visible = True
    word_app.Application.ScreenUpdating = True
    Dim word_doc As Word.Document
    Set word_doc = word_app.Documents.Open("C:\Users\tpwit\Desktop\YNH003-970-9710-GD-RT-0618.docx")
    
    Dim table_1 As Table, table_2 As Table, table_x As Table
    Dim rng As Cell
    Dim i, j
    
    Set table_1 = word_doc.Tables(1)
   
   Debug.Print table_1.Cell(0, 0).Range.Text
    
    
    
'    Debug.Print table_1.Range.Cells(2).Range.Text
'    For Each rng In table_1.Range.Cells
'        If InStr(rng.Range.Text, "��1ҳ ��1ҳ") > 0 Then
'            Debug.Print "found"
'        End If
'    Next rng
'
'    For Each rng In table_2.Range.Cells
'        If InStr(rng.Range.Text, "��2ҳ ��2ҳ") > 0 Then
'            Debug.Print "found"
'        End If
'    Next rng
'
    word_doc.Close
    Set word_app = Nothing
    
    MsgBox "Done!"

End Sub

Sub error_test()

    Dim aCell As Cell
    Dim aTable As Table
    Dim iRows As Integer
    For Each aTable In ActiveDocument.Tables
        With aTable
            For Each aCell In .Range.Cells
                With aCell
                    .Select
                    ' ��ѡ��Ԫ��������
                    Selection.SelectRow
                    ' ���ѡ������������1����õ�Ԫ��������ϲ��˶���
                    If Selection.Rows.Count > 1 Then
                        iRows = Selection.Rows.Count
                        ' ��ѡ��õ�Ԫ�������У����Եõ��õ�Ԫ�����ϲ�������
                        .Select
                        Selection.SelectColumn
                        MsgBox "RowIndex=" & .RowIndex & vbCrLf & "ColumnIndex=" & _
                                .ColumnIndex & vbCrLf & "�˵�Ԫ���Ǻϲ���Ԫ�񣬺ϲ�����Ϊ��" & iRows & _
                                "���ϲ�����Ϊ��" & Selection.Columns.Count
                                
                    Else ' ���ѡ��������Ϊ1����õ�Ԫ��������û�кϲ�
                        .Select
                        ' ��ѡ��Ԫ��������
                        Selection.SelectColumn
                        ' ���ѡ������������1����õ�Ԫ���ں���ϲ��˶���
                        If Selection.Columns.Count > 1 Then
                            MsgBox "RowIndex=" & .RowIndex & vbCrLf & "ColumnIndex=" & _
                                .ColumnIndex & vbCrLf & "�˵�Ԫ���Ǻϲ���Ԫ�񣬺ϲ�����Ϊ1��" & _
                                "�ϲ�����Ϊ��" & Selection.Columns.Count
                        Else ' ���ѡ��������Ϊ1����õ�Ԫ���ں���Ҳû�кϲ����õ�Ԫ���Ǻϲ���Ԫ��
                            MsgBox "RowIndex=" & .RowIndex & vbCrLf & "ColumnIndex=" & _
                                .ColumnIndex & vbCrLf & "�˵�Ԫ���Ǻϲ���Ԫ��"
                        End If
                    End If
                End With
            Next
        End With
    Next

End Sub

Sub gramm()

    Dim i As Integer, j As Integer
    
    For i = 1 To 5
        For j = 1 To 10
            Debug.Print j
            If j > 2 Then Exit For
        Next j
    Next i

End Sub
