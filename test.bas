Attribute VB_Name = "test"
Option Explicit

Sub test_progressbar()

    ActiveSheet.Unprotect
    Call Load_Messagebox(ActiveSheet, Range("I9"), "question", "运行程序前，请您检查数据完整性？")
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Do While Sheets("bg_paras").Range("B3") = ""
        DoEvents
    Loop
    If Sheets("bg_paras").Range("B3") = "no" Then
        ActiveSheet.Unprotect
        Call Load_Messagebox(ActiveSheet, Range("I9"), "information", "您终止了程序的运行")
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        Exit Sub
    End If
    
    Call init_pbar(ActiveSheet, Range("I9"), "测试窗体 | for test", "准备就绪。 Ready... ...", 0, 100)
    
    Call Func_Sleep(1000)
    
    Dim i As Integer
    
    For i = 1 To 100
        Call set_pbar(ActiveSheet, " you shall not pass", "程序正在运行中。。。 。。。", i, 100)
        Call Func_Sleep(25)
    Next i
    
    Call unload_pbar(ActiveSheet)
    
    Call Load_Messagebox(ActiveSheet, Range("I9"), "information", "搞定，请您检查数据！")
    
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
'        If InStr(rng.Range.Text, "第1页 共1页") > 0 Then
'            Debug.Print "found"
'        End If
'    Next rng
'
'    For Each rng In table_2.Range.Cells
'        If InStr(rng.Range.Text, "第2页 共2页") > 0 Then
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
                    ' 先选择单元格所在行
                    Selection.SelectRow
                    ' 如果选区总行数大于1，则该单元格在纵向合并了多行
                    If Selection.Rows.Count > 1 Then
                        iRows = Selection.Rows.Count
                        ' 再选择该单元格所在列，可以得到该单元格横向合并的列数
                        .Select
                        Selection.SelectColumn
                        MsgBox "RowIndex=" & .RowIndex & vbCrLf & "ColumnIndex=" & _
                                .ColumnIndex & vbCrLf & "此单元格是合并单元格，合并行数为：" & iRows & _
                                "；合并列数为：" & Selection.Columns.Count
                                
                    Else ' 如果选区总行数为1，则该单元格在纵向没有合并
                        .Select
                        ' 再选择单元格所在列
                        Selection.SelectColumn
                        ' 如果选区总列数大于1，则该单元格在横向合并了多列
                        If Selection.Columns.Count > 1 Then
                            MsgBox "RowIndex=" & .RowIndex & vbCrLf & "ColumnIndex=" & _
                                .ColumnIndex & vbCrLf & "此单元格是合并单元格，合并行数为1；" & _
                                "合并列数为：" & Selection.Columns.Count
                        Else ' 如果选区总列数为1，则该单元格在横向也没有合并，该单元格不是合并单元格
                            MsgBox "RowIndex=" & .RowIndex & vbCrLf & "ColumnIndex=" & _
                                .ColumnIndex & vbCrLf & "此单元格不是合并单元格"
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
