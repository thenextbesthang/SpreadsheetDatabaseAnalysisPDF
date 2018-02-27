Attribute VB_Name = "ExtractPDFFieldNames"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''
'Written by Catalin Bombea, 2017-02-13
'http://www.excel-first.com
'''''''''''''''''''''''''''''''''''''''''''''''

'this subroutine takes a selected file and populates a table on worksheet "templates list" with the name and path of the file
'that is used as the template
Sub UpdateTemplatesList()

'set up variables
Dim FileName As String, FilePath As String, wks As Worksheet, FolderName As String, Templ As ListObject, NewRow As Range
Dim SelectFile As FileDialog, ExitLine As Label

'set up table
Set wks = ThisWorkbook.Worksheets("Templates List")
Set Templ = wks.ListObjects(1)

'check for no data
If Not Templ.DataBodyRange Is Nothing Then Templ.DataBodyRange.Delete
'
FolderName = ThisWorkbook.Path & "\Templates"


'make sure the folder with the template is located in the same folder as this workbook
If Len(Dir$(FolderName, vbDirectory)) = 0 Then
    MsgBox "The Templates folder with your PDF forms is not in the same folder with this workbook!", vbCritical, "Templates Folder not found!"
    GoTo ExitLine
End If


'method for choosing file
Set SelectFile = Application.FileDialog(msoFileDialogFilePicker)

SelectFile.AllowMultiSelect = False
SelectFile.Title = "Select a file"
SelectFile.InitialFileName = FolderName
SelectFile.Filters.Clear
SelectFile.Filters.Add "PDF files", "*.pdf"

If Not SelectFile.Show = -1 Then GoTo ExitLine

'set what the filepath and file names are to be
'file path is in cell b2
'file name is in cell a2
FilePath = SelectFile.SelectedItems(1)
FileName = Right(FilePath, Len(FilePath) - InStrRev(FilePath, Application.PathSeparator))

'add it to the table at the bottom
Templ.ListRows.Add
Set NewRow = Templ.ListRows(Templ.ListRows.Count).Range
NewRow.Cells(1, 1) = Left(FileName, Len(FileName) - 4)
NewRow.Cells(1, 2) = FilePath

TestSheet Left(FileName, Len(FileName) - 4), FilePath

ExitLine:
Set wks = Nothing
Set SelectFile = Nothing
Set NewRow = Nothing
Set Templ = Nothing
End Sub

'this subroutine tests whether the target sheet can be mapped

Sub TestSheet(ByVal FileName As String, ByVal FilePath As String)

Dim DestSheet As Worksheet

'validates whether the file name is in the proper format
FileName = ValidateName(FileName, True)

On Error Resume Next

'sets the desination sheet
Set DestSheet = ThisWorkbook.Worksheets(FileName)

'error 9 = subscript out of range

If Err.Number = 9 Then
    'move to extract the fields
    If MsgBox("The PDF template: " & FileName & " does not have a sheet with field mappings." & vbNewLine & _
                          "Do you want to create it and extract the list of fields?", vbYesNo) = vbYes Then _
                                                ExtractFields FileName, FilePath
    
    Else
        If MsgBox("Do you want to update the list of fields?" & vbNewLine & _
                      "If ''Yes'', the existing mapping will be deleted, you have to redo the mapping.", vbYesNo) = vbYes Then _
                                            ExtractFields FileName, FilePath
End If
On Error GoTo 0

End Sub

'filename is name of template, derived from template page
'filepath is path of template, derived from template page

'sets up a worksheet in this worksheet with values based on the mapped values of the PDF
Sub ExtractFields(FileName As String, FilePath As String)

'set up variables
Dim FieldsDict As Variant, Fld As Variant
Dim FSO As Object, Dest As Worksheet, DestTbl As ListObject, NewRow As Range

Set FSO = CreateObject("Scripting.FileSystemObject")

'set up the dictionary which stores the form fields in the PDF
Set FieldsDict = ListPDFFormFields(FilePath)

'validate filename is legit
FileName = ValidateName(FileName, True)
Err.Clear
On Error Resume Next

'set the worksheet name
Set Dest = ThisWorkbook.Worksheets(FileName)

'error number 9 = subscript out of range
'so if the program is trying to set something to out of range, run this if statement
If Err.Number = 9 Then

    'set up the new worksheet that's going to hold the mapped values of the PDF; connects the data page with the PDF
    Set Dest = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    Dest.Name = FileName
    Set Dest = ThisWorkbook.Worksheets(FileName)
    
    'create table for fields
    Dest.Range("$A$1:$D$1") = Array("Source: Responses Sheets Headers", "Destination: PDF Form Field", "User Name", "Field Type")
    Dest.ListObjects.Add(xlSrcRange, Dest.Range("$A$1:$D$2"), , xlYes).Name = FileName
    
    'sets up data validation for the table on the mapping page
    'documentation found here: https://msdn.microsoft.com/en-us/library/office/ff840078.aspx
    'xlBetween = compares 2 formulas
    With Dest.Range("$A$2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=DataHeaders"
    End With
End If

On Error GoTo 0

'set the table up in the new mapping page
Set DestTbl = Dest.ListObjects(FileName)

'some error checking; if something is there in the databodyrange, delete it
If Not DestTbl.DataBodyRange Is Nothing Then DestTbl.DataBodyRange.Delete

'fieldsdict has the list of form fields in the template PDF
For Each Fld In FieldsDict.keys
    DestTbl.ListRows.Add
    Set NewRow = DestTbl.ListRows(DestTbl.ListRows.Count).Range
    'populate fields in mapping worksheet
    NewRow.Cells(1, 2) = Fld
    NewRow.Cells(1, 3) = Split(FieldsDict(Fld), "|")(0)
    NewRow.Cells(1, 4) = Split(FieldsDict(Fld), "|")(1)
Next

'resize
DestTbl.Range.Columns.AutoFit
If DestTbl.DataBodyRange Is Nothing Then
    MsgBox "No field found, the form might have been created with Live Cycle Designer, this type of form cannot be filled with this tool."
End If

Set FieldsDict = Nothing
End Sub

Function ListPDFFormFields(FilePath As String) As Variant
Set ListPDFFormFields = CreateObject("Scripting.Dictionary")
    Dim AcroExchAVDoc As Object, objAcroPDDoc As Object, objJSO As Object
    Dim AcroExchApp As Object
    Dim AFORMAUT As Object 'AFORMAUTLib.AFormApp
    Dim FormField As Variant 'AFORMAUTLib.Field
    Dim FormFields As Variant 'AFORMAUTLib.Fields
    Dim bOK As Boolean
    Dim sFields As String
    Dim sTypes As String
    Dim sFieldName As String
   
    On Error GoTo ErrorHandler
     
    Set AcroExchApp = CreateObject("AcroExch.App")
    Set AcroExchAVDoc = CreateObject("AcroExch.AVDoc")
    
    bOK = AcroExchAVDoc.Open(FilePath, "")
    AcroExchAVDoc.BringToFront
    AcroExchApp.Hide
    
    If (bOK) Then
        Set objAcroPDDoc = AcroExchAVDoc.GetPDDoc
        Set objJSO = objAcroPDDoc.GetJSObject
      
        Set AFORMAUT = CreateObject("AFormAut.App")
        Set FormFields = AFORMAUT.Fields
        
        For Each FormField In FormFields
           If FormField.IsTerminal Then ListPDFFormFields.Add FormField.Name, objJSO.getField(FormField.Name).UserName & "|" & FormField.Type
        Next FormField
        
        AcroExchAVDoc.Close True
    End If
    
    Set AcroExchAVDoc = Nothing
    Set AcroExchApp = Nothing
    Set AFORMAUT = Nothing
    Exit Function
         
ErrorHandler:
MsgBox "FieldList Error: " + Str(Err.Number) + " " + Err.Description + " " + Err.Source
    
End Function
