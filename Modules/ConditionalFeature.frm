VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConditionalFeature 
   Caption         =   "ConditionalFeature"
   ClientHeight    =   9570.001
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7575
   OleObjectBlob   =   "ConditionalFeature.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConditionalFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------Form Submit Button------

Private Sub BuildFormulasButton_Click()
    Dim partWS As Worksheet
    Set partWS = Worksheets("PartLib Table")

    On Error GoTo convErr

    'Set the static or variable LTol
    If Me.ComboBox9 <> vbNullString Then
        Call Worksheets("PartLib Table").ApplyFormula(targetCol:="Lower", varCol:=Me.ComboBox9)
    ElseIf Me.LowerTextBox <> vbNullString Then
        If IsNumeric(Me.LowerTextBox) Then
            Call Worksheets("PartLib Table").ApplyFormula(targetCol:="Lower", limit:=Me.LowerTextBox)
        Else
            Err.Raise Number:=vbObjectError + 1000, Description:="Lower value"
        End If
    End If
    'Set the static or variable Nominal
    If Me.ComboBox10 <> vbNullString Then
        Call Worksheets("PartLib Table").ApplyFormula(targetCol:="Nominal", varCol:=Me.ComboBox10)
    ElseIf Me.NominalTextBox <> vbNullString Then
        If IsNumeric(Me.NominalTextBox) Then
            Call Worksheets("PartLib Table").ApplyFormula(targetCol:="Nominal", limit:=Me.NominalTextBox)
        Else
            Err.Raise Number:=vbObjectError + 1000, Description:="Nominal value"
        End If
    End If
    'Set the static or variable UTol
    If Me.ComboBox11 <> vbNullString Then
        Call Worksheets("PartLib Table").ApplyFormula(targetCol:="Upper", varCol:=Me.ComboBox11)
    ElseIf Me.UpperTextBox <> vbNullString Then
        If IsNumeric(Me.UpperTextBox) Then
            Call Worksheets("PartLib Table").ApplyFormula(targetCol:="Upper", limit:=Me.UpperTextBox)
        Else
            Err.Raise Number:=vbObjectError + 1000, Description:="Upper value"
        End If
    End If
    'If a feature it checked off and we have a selection in the adjacent comboBox, then construct the variable formula
    For i = 1 To 8
        If Me.Controls("CheckBox" & i).Value = True And Me.Controls("ComboBox" & i) <> vbNullString Then
            Call Worksheets("PartLib Table").ApplyFormula(targetCol:=Me.Controls("ColLabel" & i), varCol:=Me.Controls("ComboBox" & i))
        End If
    Next i
    
    
    Unload Me
    Exit Sub

convErr:
    MsgBox "Could not convert" & vbCrLf & Err.Description, vbCritical
    Unload Me
End Sub











'-----Checkboxes------

Private Sub Checkbox1_Click()
    Me.Controls("ComboBox1").Enabled = Me.Controls("Checkbox1").Value
End Sub
Private Sub Checkbox2_Click()
    Me.Controls("ComboBox2").Enabled = Me.Controls("Checkbox2").Value
End Sub
Private Sub Checkbox3_Click()
    Me.Controls("ComboBox3").Enabled = Me.Controls("Checkbox3").Value
End Sub
Private Sub Checkbox4_Click()
    Me.Controls("ComboBox4").Enabled = Me.Controls("Checkbox4").Value
End Sub
Private Sub Checkbox5_Click()
    Me.Controls("ComboBox5").Enabled = Me.Controls("Checkbox5").Value
End Sub
Private Sub Checkbox6_Click()
    Me.Controls("ComboBox6").Enabled = Me.Controls("Checkbox6").Value
End Sub
Private Sub Checkbox7_Click()
    Me.Controls("ComboBox7").Enabled = Me.Controls("Checkbox7").Value
End Sub
Private Sub Checkbox8_Click()
    Me.Controls("ComboBox8").Enabled = Me.Controls("Checkbox8").Value
End Sub



'----Lower-----

Private Sub ComboBox9_Change()
    If Me.ComboBox9.Value = "" Then Exit Sub
    Me.LowerTextBox.Value = ""
    
End Sub


Private Sub LowerTextBox_Change()
    If Me.LowerTextBox.Value = "" Then Exit Sub
    Me.ComboBox9.Value = ""
End Sub


'----Nominal-----

Private Sub ComboBox10_Change()
    If Me.ComboBox10.Value = "" Then Exit Sub
    Me.NominalTextBox.Value = ""
End Sub

Private Sub NominalTextBox_Change()
    If Me.NominalTextBox.Value = "" Then Exit Sub
    Me.ComboBox10.Value = ""
End Sub

'----Upper-----

Private Sub ComboBox11_Change()
    If Me.ComboBox11.Value = "" Then Exit Sub
    Me.UpperTextBox.Value = ""
End Sub

Private Sub UpperTextBox_Change()
    If Me.UpperTextBox.Value = "" Then Exit Sub
    Me.ComboBox11.Value = ""
End Sub

Private Sub UserForm_Click()

End Sub
