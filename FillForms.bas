Attribute VB_Name = "FillForms"
Option Explicit

Dim totalHoursInExt As Long
Dim extStartDate As Date
Dim lastBillableDate As Date
Dim hoursPerDay As Long
Dim hoursColumn As Long
Dim dateLastApproved As Date
Dim dateLastWritten As Date
Dim startDate As Date
Dim extensionSheet As Worksheet
Dim neededExtensions As Long

'preliminary subroutine, calls writepdfforms
'called from the double click method
'shName = worksheet that gets the double click
'RowNumber = row of the double clicked cell

Public Sub FillSelectedForms(ShName As Worksheet, RowNumber As Long)

Dim cell As Range, wks As Worksheet, Templ As ListObject, ExitLine As Label

        Dim i As Long
        Set extensionSheet = ThisWorkbook.Worksheets("Extensions")
        
'get template list
Set wks = ThisWorkbook.Worksheets("Templates List")
Set Templ = wks.ListObjects(1)

If Templ.ListColumns(1).DataBodyRange Is Nothing Then
    MsgBox "No data found in Templates List", vbInformation, "Missing Data"
    GoTo ExitLine
End If

'databodyrange = first column in the data (not header) cell 1
Set cell = Templ.ListColumns(1).DataBodyRange.Cells(1)
      
      'hours in extension = 200
      'start date
      'end date
      totalHoursInExt = 200
     On Error GoTo errhandlr
     
      hoursPerDay = (extensionSheet.Cells(RowNumber, 35) / 7)
      
      
      If (hoursPerDay > 0) Then
            extStartDate = DateAdd("d", 200 / hoursPerDay, extensionSheet.Cells(RowNumber, 8))
      Else
            
      End If
      
      
      lastBillableDate = DateAdd("d", 365, extensionSheet.Cells(RowNumber, 36))
      
      neededExtensions = extensionSheet.Cells(RowNumber, 7)
            
           'do while
           'two conditions to make sure for here
           'extStartDate never exceeds lastBillableDate
     MsgBox (extStartDate)
     
           Do While extStartDate < lastBillableDate And neededExtensions > 0
                'every extension will be for 200 hours
                extensionSheet.Cells(RowNumber, 82) = extStartDate
                WritePDFForms ShName.Name, RowNumber, cell, cell.Offset(0, 1)
                'update last extension
                extStartDate = DateAdd("d", 200 / hoursPerDay, extStartDate)
                neededExtensions = neededExtensions - 1
            Loop

ExitLine:
Set Templ = Nothing
Set wks = Nothing
Set cell = Nothing

errhandlr:
  '  MsgBox (RowNumber)
    

End Sub

'wksname = sheet name where double click is initialized
'rw = row number of double click
'fname = template file name
'strPDFPath = template file path and name
Sub WritePDFForms(WksName As String, Rw As Long, ByVal Fname As String, ByVal StrPDFPath As String)

'set up variables
Dim MapDict As Variant, ColDict As Variant, TypeDict As Variant, MapKey As Variant, ColKey As Variant, FolderPath As String
Dim LastRow As Long, objAcroApp As Object, objAcroAVDoc As Object, MapSheet As String, wks As Worksheet
Dim objAcroPDDoc As Object, objJSO As Object, strPDFOutPath As String, Fld As String, ResponseText As String
Dim cell As Range, FieldsNotMapped As String, LastCol As Integer, FSO As Object, ErrExit As Label
Dim ClearForm, check As Boolean: check = False

'set objects and the file systems; used for navigating windows environment
'wks = worksheet object of the double click initializer
Set wks = ThisWorkbook.Worksheets(WksName)
Set FSO = CreateObject("Scripting.FileSystemObject")

'finds the last column on the datapage that initialized the double click
LastCol = wks.Cells.Find("*", wks.Cells(1, 1), , , xlByColumns, xlPrevious).column
'Application.ScreenUpdating = False

Set ColDict = CreateObject("Scripting.Dictionary")
Set MapDict = CreateObject("Scripting.Dictionary")
MapSheet = ValidateName("Request_to_bill_additional_sem", True)

For Each cell In ThisWorkbook.Worksheets(MapSheet).ListObjects(MapSheet).ListColumns(1).DataBodyRange
    'if the cell in the first and second column is empty
    If Len(cell) > 0 And Len(cell.Offset(0, 1)) > 0 Then
        'if the value hasn't been added to the map dictionary then add both the second column and the first column values
        If MapDict.exists(cell.Offset(0, 1).Text) = False Then

            'map looks like
            'field value on PDF.....field value on data sheet
            MapDict.Add cell.Offset(0, 1).Text, cell.Text
        Else
        End If
    Else
        FieldsNotMapped = FieldsNotMapped & vbNewLine & cell.Offset(0, 1) 'these fields will not be filled!
    End If
Next cell

'error checking for template
If FSO.FileExists(StrPDFPath) = False Then
    MsgBox "The file named ''" & Fname & ".pdf'' was not found in the PDF Templates folder." & _
            vbNewLine & "" & vbNewLine & _
            ""
    GoTo ErrExit
End If

On Error Resume Next
    'Create Acrobat objects (late binding, no need to set a Reference to Acrobat library).
    'first create the general level acrobat application - ArcoExch.App controls the appearance of Acrobat, size of Application Windows
    Set objAcroApp = CreateObject("AcroExch.App")
    If Err.Number <> 0 Then MsgBox "Could not create the AcroExch.App object!", vbCritical, "Object error": GoTo ErrExit
    'second create the window containing an open pdf document; use this object to select text, find text, or print pages
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")
    If Err.Number <> 0 Then MsgBox "Could not create the AcroExch.AVDoc object!", vbCritical, "Object error": GoTo ErrExit
On Error GoTo 0

'Open the PDF file.
If objAcroAVDoc.Open(StrPDFPath, "") = True Then
    
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc
    Set objJSO = objAcroPDDoc.GetJSObject
    ClearForm = objJSO.ResetForm()  'clear the form

    
    'wks = double click initializing sheet
       ' On Error Resume Next
        For Each cell In wks.ListObjects(1).HeaderRowRange
            'map the data in the selected row to the header in that row; used for finding specific data
            If Len(wks.Cells(Rw, cell.column).Text) > 0 Then
                'key, item
                ColDict.Add cell.Text, wks.Cells(Rw, cell.column).Text
                
                'map looks like this
                'consumer name, 1
                'last name, 2
                'first name, 3
                '...
                'email, 78
                
                
                
            End If
        Next cell
        
        On Error Resume Next

        For Each MapKey In MapDict.keys
            If ColDict.exists(MapDict(MapKey)) Then
                    If objJSO.getField(MapKey).Type = "text" Then
                        objJSO.getField(MapKey).Value = ColDict(MapDict(MapKey))
                    ElseIf objJSO.getField(MapKey).Type = "checkbox" Then
                        If ColDict(MapDict(MapKey)) = True Then
                            objJSO.getField(MapKey).Value = "Yes"
                        Else
                            objJSO.getField(MapKey).Value = "No"
                        End If
                    Else
                    End If
            End If
        Next
               

          
        'set up the folder path
        FolderPath = ThisWorkbook.Path & "\Extensions\" & Year(extStartDate)
        
        'if the folder path does not exist, make it
        If FSO.FolderExists(FolderPath) = False Then MkDir FolderPath
                  
        'rename the folder path with a validated name
        strPDFOutPath = FolderPath & "\" & wks.Cells(Rw, 1) & "." & "200hours." & Month(extStartDate) & "." & Day(extStartDate) & "." & Year(extStartDate) & ".pdf"

        'check if file was saved in adobe
        If objAcroPDDoc.Save(1, strPDFOutPath) <> True Then
      '      MsgBox "File was not saved."
        Else
           'success!

        End If


        'close adobe
        objAcroAVDoc.Close True

Else

    MsgBox "Could not open the file!", vbCritical, "File error"


End If

    objAcroAVDoc.Close True

ErrExit:
Application.ScreenUpdating = True
On Error Resume Next
objAcroApp.Exit
Set objJSO = Nothing
Set objAcroPDDoc = Nothing
Set objAcroAVDoc = Nothing
Set objAcroApp = Nothing
Set MapDict = Nothing
Set ColDict = Nothing
Set TypeDict = Nothing
Set FSO = Nothing
On Error GoTo 0

End Sub

'validates file names to ensure integrity
Function ValidateName(ByVal TextString As String, Optional IsSheetName As Boolean = False) As String
Dim CurrChar, TheCleanString As String, ChrPos As Integer
   TheCleanString = ""
   CurrChar = ""
   
   Const ValidChars = "[A-Z,a-z,0-9, ,_,.,-]"
   
     For ChrPos = 1 To Len(TextString)
      CurrChar = Mid(TextString, ChrPos, 1)
      
      If CurrChar Like ValidChars Then
         TheCleanString = TheCleanString & CurrChar
      Else
      TheCleanString = TheCleanString & "_"
      End If
   Next
   
   ValidateName = StrConv(TheCleanString, vbProperCase)
    
   If IsSheetName Then ValidateName = Replace(Replace(Left(Replace(ValidateName, " ", "_"), 30), "__", "_"), "__", "_")
   
End Function




