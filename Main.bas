Attribute VB_Name = "Main"
Option Explicit

Dim dic_files As Object

Sub about()

    Call Load_Messagebox(sht_main, Range("I9"), "information", "��������Ƽ�    ���    15809212391")
    
End Sub

Sub get_file_rec(path_dir)

    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim my_file, my_folder
    
    For Each my_file In fso.getFolder(path_dir).Files
        dic_files(my_file) = ""
    Next my_file
    
    If fso.getFolder(path_dir).subFolders.Count > 0 Then
        For Each my_folder In fso.getFolder(path_dir).subFolders
            Call get_file_rec(my_folder)
        Next my_folder
    End If

End Sub

Sub main_sub()


    ' Ĭ�����ݶ��洢�ڵ�ǰ�ļ�ͬĿ¼�����Data�ļ�����
    Dim path_root As String: path_root = ThisWorkbook.Path & "\Data"
    Dim dic_keywords As Object
    Dim var_key  As Variant
    Dim wb As Workbook, sht As Worksheet
    Dim rng_mark As Range, col_mark As Integer  ' record the column or cell of the key_word:��λ
    Dim max_row As Integer, int_row As Integer
    Dim Total_Count, current_file_index As Integer
    
    Dim sht_main As Worksheet: Set sht_main = Sheets("ͳ�Ƽ���")
    Call init_pbar(sht_main, Range("J8"), "ͳ�Ƽ���", "׼����", 0, 0)
    
    ' initilize dics
    Set dic_files = CreateObject("Scripting.Dictionary")
    Set dic_keywords = CreateObject("Scripting.Dictionary")
    
    ' get file name dictionary
    Call get_file_rec(path_root)
    
    Total_Count = 0
    ' get keywords dictionary
    For int_row = 9 To Range("I65536").End(xlUp).Row
        dic_keywords(Cells(int_row, 9).Text) = ""
    Next int_row
    
    Call set_pbar(sht_main, "ͳ�Ƽ���", "��ʼͳ�Ƽ���", 0, dic_files.Count)
    For Each var_key In dic_files.keys
        
        current_file_index = current_file_index + 1
        Call set_pbar(sht_main, _
                        "ͳ�Ƽ���", "ͳ��:" & _
                        Split(var_key, "\")(UBound(Split(var_key, "\"))) & _
                        "    �ѻ�ȡ��" & Total_Count & "��", _
                        current_file_index, dic_files.Count)
        
        Application.ScreenUpdating = False
        Set wb = Workbooks.Open(var_key)
        For Each sht In wb.Sheets
        
            col_mark = 0
            ' ���ȫ�Ŀո� ���е�
            sht.Cells.Replace what:=" ", Replacement:=""
'            sht.Cells.Replace what:=Chr(10), Replacement:=""
'            sht.Cells.Replace what:=Chr(13), Replacement:=""
'
           ' get marked column or cell
            On Error Resume Next
            Set rng_mark = sht.Cells.Find("������λ", lookat:=xlWhole)
            If rng_mark Is Nothing Then
                Set rng_mark = sht.Cells.Find("��λ", lookat:=xlWhole)
            End If
            If Not rng_mark Is Nothing Then
                col_mark = rng_mark.Column
            Else
                col_mark = 0
            End If
            On Error GoTo 0
            
            If col_mark = 0 Then Exit For
            
            max_row = sht.Cells(Rows.Count, col_mark).End(xlUp).Row
            
            For int_row = 1 To max_row
                If sht.Cells(int_row, col_mark) <> "" And dic_keywords.Exists(sht.Cells(int_row, col_mark).Text) Then
                    Total_Count = Total_Count + 1
                End If
            Next int_row
            
        Next sht
        wb.Close SaveChanges:=False
        
        Application.ScreenUpdating = True
        sht_main.Range("L2") = Total_Count
        
    Next var_key
    
    Application.ScreenUpdating = True
    Call unload_pbar(sht_main)
    Call Load_Messagebox(sht_main, Range("J8"), "information", "��������ɣ�����ȡ������Ϊ:" & CStr(Total_Count))

End Sub

