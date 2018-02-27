VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub runExtensions()

    Dim i As Long
    Dim rowCount As Long
    
    rowCount = ActiveSheet.Range("A1").End(xlDown).Row
    
    
    For i = 2 To 10 'rowCount
        FillSelectedForms ActiveSheet, i
    Next i
    
    
End Sub
