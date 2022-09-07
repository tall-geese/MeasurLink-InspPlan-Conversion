VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HideFeatureCond 
   Caption         =   "Hide Feature(s) Conditionally"
   ClientHeight    =   4680
   ClientLeft      =   -240
   ClientTop       =   -1050
   ClientWidth     =   7185
   OleObjectBlob   =   "HideFeatureCond.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HideFeatureCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'*************************************************************
'*                  HideFeatureCond
'*
'*************************************************************
'*************************************************************



Private Sub BuildHiddenFormulasButton_Click()
    Dim formula As String
    
    'If the User entered something in the PartNumber Text field
    If Me.PartNumberTextBox.Value <> vbNullString Then
        On Error GoTo varErr
        
        Dim partField As String
        partField = Replace(Me.PartNumberTextBox.Value, " ", "") 'Ignore spaces
        
        Dim partNumCol As Collection
        Set partNumCol = New Collection
        Dim seperations() As String
        
        If InStr(partField, ",") > 0 Then 'Break up any comma-seperated values
            seperations = Split(partField, ",")
        End If
        
        If ((Not seperations) = -1) Then 'If nothing was comma seperated, then create the array and set the value to the partNumber field
            ReDim Preserve seperations(0)
            seperations(0) = partField
        End If
        For i = 0 To UBound(seperations)
            If InStr(seperations(i), "-") > 0 Then 'If we found a Range of values, split them up and do the high-low math
                Dim partRange() As String
                partRange = Split(seperations(i), "-") 'iterate over the difference, set lowNum + current iteration
                If UBound(partRange) <> 1 Then
                    'There should be exactly to items: from and to
                    Err.Raise Number:=vbObjectError + 1000, Description:="There should be exactly 2 items per '-' Delimiter"
                End If
                
                Dim prefix As String
                Dim fromNum As String
                Dim toNum As String
                
                fromNum = ThisWorkbook.GetNumericSuffix(partRange(0))
                toNum = ThisWorkbook.GetNumericSuffix(partRange(1))
                prefix = Split(partRange(0), fromNum)(0)
                
                If prefix <> vbNullString Then
                    If prefix <> Split(partRange(1), toNum)(0) Then 'The leading values arent the same
                        Err.Raise Number:=vbObjectError + 1000, Description:="The leading characters of the part numbers between " _
                            & "the range are not the same"
                    End If
                    partRange(0) = fromNum
                    partRange(1) = toNum
                End If
                
                
                If CDbl(partRange(1)) < CDbl(partRange(0)) Then
                    Err.Raise Number:=vbObjectError + 1000, Description:="End of range is smaller than the beginning"
                End If
                
                For j = 0 To CDbl(partRange(1)) - CDbl(partRange(0))
                    partNumCol.Add (prefix & CStr(CDbl(partRange(0) + j)))
                Next j
            Else
                partNumCol.Add (seperations(i)) 'Otherwise, just add the number
            End If
        Next i
        
        'Build the formula
        If partNumCol.Count = 1 Then
                'Output of...
                '=IF('START HERE'!$C$8=1642652,"",
            formula = "=IF('START HERE'!$C$8=" & Chr(34) & partNumCol.item(1) & Chr(34) & "," & Chr(34) & Chr(34) & ","
        ElseIf partNumCol.Count > 1 Then
                'Output of....
                '=IF(OR('START HERE'!$C$8=1642652,'START HERE'!$C$8=1642653),"",
            formula = "=IF(OR("
            For Each partNum In partNumCol
                formula = formula & "'START HERE'!$C$8=" & Chr(34) & partNum & Chr(34)
                If Not (partNumCol.item(partNumCol.Count) = partNum) Then
                    formula = formula & ","
                End If
            Next partNum
            formula = formula & ")," & Chr(34) & Chr(34) & ","
        End If
        
    ElseIf Me.VariableComboBox <> vbNullString And Me.VariableTextBox <> vbNullString Then
        Dim colNum As Integer
        colNum = Worksheets("Variables").GetCol(Me.VariableComboBox.Value)
            'Output like....
            '=IF(VLOOKUP('START HERE'!$C$8,Variables!$A$2:$AZ$500,22,FALSE)="cerv","",VLOOKUP('START HERE'!$C$8,Variables!$A$2:$AZ$500,5,FALSE))
        formula = "=IF(VLOOKUP('START HERE'!$C$8,Variables!$A$2:$AZ$500," & colNum & ",FALSE)=" & Chr(34) & Me.VariableTextBox & Chr(34) & "," & Chr(34) & Chr(34) & ","

    End If
    
    'TODO: grab the collection of cells from our tag and take the formula and pass them to PartLib routine to handle it
    Call Worksheets("PartLib Table").SetHiding(Me.Tag, formula)
    
10
    Unload Me
    
    Exit Sub
varErr:
    If Err.Number <> vbObjectError + 1000 Then
       MsgBox "Couldn't convert the Part Number(s), Try again or try setting a variable instead", vbCritical
    Else
        MsgBox Err.Description, vbCritical
    End If
    On Error GoTo 0
    GoTo 10
End Sub






Private Sub PartNumberTextBox_Change()
    If Me.PartNumberTextBox.Value = vbNullString Then Exit Sub
    Me.VariableComboBox.Value = vbNullString
    Me.VariableTextBox.Value = vbNullString
End Sub


Private Sub UserForm_Click()

End Sub

Private Sub VariableComboBox_Change()
    If Me.VariableComboBox.Value = vbNullString Then Exit Sub
    Me.PartNumberTextBox.Value = vbNullString
End Sub


Private Sub VariableTextBox_Change()
    If Me.VariableTextBox.Value = vbNullString Then Exit Sub
    Me.PartNumberTextBox.Value = vbNullString
        
End Sub


