VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HideFeatureCond 
   Caption         =   "Hide Feature(s) Conditionally"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   OleObjectBlob   =   "HideFeatureCond.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HideFeatureCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BuildHiddenFormulasButton_Click()
    Dim formula As String
    
    'If the User entered something in the PartNumber Text field
    If Me.PartNumberTextBox.Value <> vbNullString Then
        On Error GoTo convertErr
        
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
                If CDbl(partRange(1)) < CDbl(partRange(0)) Then GoTo convertErr
                
                For j = 0 To CDbl(partRange(1)) - CDbl(partRange(0))
                    partNumCol.Add (CStr(CDbl(partRange(0) + j)))
                Next j
            Else
                partNumCol.Add (seperations(i)) 'Otherwise, just add the number
            End If
        Next i
        
        'Build the formula
        If partNumCol.Count = 1 Then
                'Output of...
                '=IF('START HERE'!$C$8=1642652,"",
            formula = "=IF('START HERE'!$C$8=" & partNumCol.Item(1) & "," & Chr(34) & Chr(34) & ","
        ElseIf partNumCol.Count > 1 Then
                'Output of....
                '=IF(OR('START HERE'!$C$8=1642652,'START HERE'!$C$8=1642653),"",
            formula = "=IF(OR("
            For Each partNum In partNumCol
                formula = formula & "'START HERE'!$C$8=" & partNum
                If Not (partNumCol.Item(partNumCol.Count) = partNum) Then
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
    
convertErr:
    MsgBox "Couldn't convert the Part Number(s), Try again or try setting a variable instead", vbCritical
    GoTo 10
    
varErr:
    MsgBox "Couldn't create a formula with these values", vbCritical
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
