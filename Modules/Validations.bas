Attribute VB_Name = "Validations"
Public valWB As Workbook


Public Sub OpenDataValidations(Optional pass As Variant, Optional readMode As Variant)
    If valWB Is Nothing Then
        On Error GoTo wbOpenErr
        If Not (IsMissing(pass) And IsMissing(readMode)) Then 'Called in this manner by InsertNewValidation
            Set valWB = Workbooks.Open(Filename:=DataSources.DATA_VALIDATION_PATH, UpdateLinks:=0, ReadOnly:=readMode, Password:=pass, WriteResPassword:=pass)
        Else
            Set valWB = Workbooks.Open(Filename:=DataSources.DATA_VALIDATION_PATH, UpdateLinks:=0, ReadOnly:=True)
        End If
    End If
    
    Exit Sub
    
wbOpenErr:
    result = MsgBox("Unable to Open Validations Workbook" & vbCrLf & "if you supplied a write Password, the password may be incorrect" & vbCrLf & "Otherwise the Routine" _
            & "MapDataValidations workbook may not exist anymore or the network may be down", vbCritical)
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseDataValidations(Optional saveWB As Boolean)
    On Error Resume Next
    Workbooks("RoutineMapDataValidations.xlsm").Close SaveChanges:=saveWB
    Set valWB = Nothing
End Sub


'***********   Called by Import New Data Validations Button  ******************
Public Sub ImportValidationValues()
    Call OpenDataValidations
    
    Dim valWS As Worksheet
    Set valWS = valWB.Worksheets("Data Validations")
    Dim commWS As Worksheet
    Set commWS = valWB.Worksheets("StandardComments")
    Dim inspWS As Worksheet
    Set inspWS = valWB.Worksheets("InspMethods")
    
    
    Dim methodAssignmentRange As Range
    Set methodAssignmentRange = valWS.Range("$E$2:$F$" & valWS.Range("F2").End(xlDown).Row)
    methodAssignmentRange.Copy Destination:=ThisWorkbook.Worksheets("Data Validations").Range("E2")
    
    Dim commentRange As Range
    Set commentRange = commWS.Range("$A$2:$A$" & commWS.Range("A2").End(xlDown).Row)
    commentRange.Copy Destination:=ThisWorkbook.Worksheets("StandardComments").Range("A2")
    
    Dim inspRange As Range
    Set inspRange = inspWS.Range("$A$2:$A$" & inspWS.Range("A2").End(xlDown).Row)
    inspRange.Copy Destination:=ThisWorkbook.Worksheets("InspMethods").Range("A2")
   
    Call CloseDataValidations

End Sub



'***********   Called by Insert Validations Button  ******************

Public Function ValidationValueExists(inputVal As String, targetCol As Integer) As Boolean
    'Check if this Comment or Inspection Method already exists
    If targetCol = 13 Then
        ValidationValueExists = valWB.Sheets("StandardComments").ValueExists(inputVal)
    ElseIf targetCol = 14 Then
        ValidationValueExists = valWB.Sheets("InspMethods").ValueExists(inputVal)
    End If
End Function

Public Sub InsertNewValidation(newVal As String, targetCol As Integer, userPass As String)
    'Add a new inspection mehtod to the Write version of the Validations Workbook
    Call CloseDataValidations
    Call OpenDataValidations(pass:=userPass, readMode:=False)
    
    If valWB Is Nothing Then Exit Sub
    If targetCol = 13 Then
        valWB.Sheets("StandardComments").InsertNewValue (newVal)
    ElseIf targetCol = 14 Then
        'TODO: prompt the user for variable/Attribue to add in the new validation.
        'Becuase this must also be set in the DataValidations sheet of the outside workbook
        Load InspectionTypeForm
        
        InspectionTypeForm.Show
               
        If InspectionTypeForm.VariableRad.Value = True Then
            valWB.Sheets("Data Validations").InsertNewValue newVal, "Variable"
        ElseIf InspectionTypeForm.AttributeRad.Value = True Then
            valWB.Sheets("Data Validations").InsertNewValue newVal, "Attribute"
        End If
        
        Unload InspectionTypeForm
    
        valWB.Sheets("InspMethods").InsertNewValue (newVal)
    End If
    
    
End Sub















'********************   Deprecated as of 2.0.0  *****************************


'*************   Called by SetValidations Button  *******************
'
'Public Sub SetDataValidations()
'    Call OpenDataValidations
'    valWB.Sheets("StandardComments").SetValReference (ThisWorkbook.Name)
'    valWB.Sheets("InspMethods").SetValReference (ThisWorkbook.Name)
'End Sub



''***********   Called by PartLib On_Change  ******************
'
'
Public Sub SetInspMethodValidation(cell As Range)
    '=INDIRECT("[RoutineMapDataValidations.xlsm]InspMethods!C2#")    [for N9]
    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="=InspMethods!$C" & cell.Row - 7 & "#"
        .ShowError = False
    End With

'=InspMethods!$C2#

End Sub
'
'
'Public Sub SetCommentsValidation(cell As Range)
'    '=INDIRECT("[RoutineMapDataValidations.xlsm]StandardComments!C2#")    [for M9]
'    If valWB Is Nothing Then Call OpenDataValidations
'    With cell.Validation
'        .Delete
'        .Add Type:=xlValidateList, Formula1:="=INDIRECT(" & Chr(34) & "[RoutineMapDataValidations.xlsm]StandardComments!C" & cell.Row - 7 & "#" & Chr(34) & ")"
'        .ShowError = False
'    End With
'
'End Sub













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
    Set sample = Worksheets("PartLib Table").Range("Z9:Z400") 'Change this row here
    Dim workingCell As Range
    For Each workingCell In sample
        With workingCell.Validation
            .Delete
            Dim tempstring As String
            'Change formula1 here
'            tempstring = "=INDIRECT(" & Chr(34) & "[RoutineMapDataValidations.xlsm]InspMethods!C" & workingCell.Row - 7 & "#" & Chr(34) & ")"
            .Add Type:=xlValidateList, Formula1:="=MachOps"
            .ShowError = False 'Change as needed if you need to set an erorr or not
        End With
    Next workingCell

End Sub





