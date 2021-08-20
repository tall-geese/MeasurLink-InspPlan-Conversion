Attribute VB_Name = "Validations"
Public valWB As Workbook


Public Sub OpenDataValidations()
    If valWB Is Nothing Then
        Set valWB = Workbooks.Open(Filename:=DataSources.DATA_VALIDATION_PATH, UpdateLinks:=0, ReadOnly:=True)
    End If
End Sub

Public Sub SetDataValidations()
    valWB.Sheets("StandardComments").SetValReference (ThisWorkbook.Name)
    valWB.Sheets("InspMethods").SetValReference (ThisWorkbook.Name)

End Sub


Public Sub SetInspMethodValidation(cell As Range)
    '=INDIRECT("[RoutineMapDataValidations.xlsm]InspMethods!C2#")    [for N9]
    If valWB Is Nothing Then Call OpenDataValidations
    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=INDIRECT(" & Chr(34) & "[RoutineMapDataValidations.xlsm]InspMethods!C" & cell.Row - 7 & "#" & Chr(34) & ")"
        .ShowError = False
    End With

End Sub


Public Sub SetCommentsValidation(cell As Range)
    '=INDIRECT("[RoutineMapDataValidations.xlsm]StandardComments!C2#")    [for M9]
    If valWB Is Nothing Then Call OpenDataValidations
    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=INDIRECT(" & Chr(34) & "[RoutineMapDataValidations.xlsm]StandardComments!C" & cell.Row - 7 & "#" & Chr(34) & ")"
        .ShowError = False
    End With

End Sub

Sub test()
    Set workingRange = Range("X4:X400")
    For Each cell In workingRange
        With cell.Validation
            '"=INDIRECT(" & chr(34) & "[RoutineMapDataValidations.xlsm]InspMethods!C2#" & chr(34) & ")"
            .Delete
            .Add Type:=xlValidateList, Formula1:="=MachineHead"
            .ShowError = False
        End With
    Next cell
End Sub










'Only using this to set that data validations once, can delete this after
Private Sub TempSetValidtion()
    Dim sample As Range
    Set sample = Range("W9:W400") 'Change this row here
    Dim workingCell As Range
    For Each workingCell In sample
        With workingCell.Validation
            .Delete
            Dim tempstring As String
            'Change formula1 here
'            tempstring = "=INDIRECT(" & Chr(34) & "[RoutineMapDataValidations.xlsm]InspMethods!C" & workingCell.Row - 7 & "#" & Chr(34) & ")"
            .Add Type:=xlValidateList, Formula1:="=AxisOffset"
            .ShowError = False 'Change as needed if you need to set an erorr or not
        End With
    Next workingCell

End Sub



