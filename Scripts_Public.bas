Attribute VB_Name = "Scripts_Public"
Option Explicit

Public Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long

Sub Func_Sleep(int_millisecond As Integer)
    ' ∫¡√Îº∂—” ±¥˙¬Î
    Dim t
    Dim int_count As Integer
    
    t = timeGetTime
    
    Do While timeGetTime - t < int_millisecond
        DoEvents
    Loop
    
End Sub
