Attribute VB_Name = "Module1"
Public Sub checkBoxControl()


    Dim cb As CheckBox
    Dim myRangeTrue As Range, myRangeFalse As Range, cel As Range
    Dim wks As Worksheet

    Set wks = Worksheets("Extensions") 'adjust sheet to your needs

    Set myRangeTrue = wks.Range("S2:S63,U2:U63, W2:W63, Y2:Y63, AJ2:AJ63, AL2:AL63, AQ2:AQ63, AQ2:AQ63, AS2:AS63, AW2:AW63, BA2:BA63") ' adjust range to your needs
    Set myRangeFalse = wks.Range("AN2:AN63, AU2:AU63, BC2:BC63") ' adjust range to your needs
    
    

    For Each cb In wks.CheckBoxes
        cb.Delete
    Next

    For Each cel In myRangeTrue
            Set cb = wks.CheckBoxes.Add(cel.Left, cel.Top, 30, 6) 'you can adjust left, top, height, width to your needs
            cb.Value = True
            With cb
                .Caption = ""
                .LinkedCell = cel.Address
            End With
    Next
    
    For Each cel In myRangeFalse
            Set cb = wks.CheckBoxes.Add(cel.Left, cel.Top, 30, 6) 'you can adjust left, top, height, width to your needs
            cb.Value = False
            With cb
                .Caption = ""
                .LinkedCell = cel.Address
            End With
    Next
    
    
    
End Sub

