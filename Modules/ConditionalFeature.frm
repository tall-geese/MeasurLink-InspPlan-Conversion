VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConditionalFeature 
   Caption         =   "ConditionalFeature"
   ClientHeight    =   8850.001
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

Private Sub NoneButton_Click()
    Me.LowerTextBox.Enabled = Not (Me.NoneButton.Value)
    Me.UpperTextBox.Enabled = Not (Me.NoneButton.Value)
    
    Me.ComboBox9.Enabled = Not (Me.NoneButton.Value)
    Me.ComboBox10.Enabled = Not (Me.NoneButton.Value)
End Sub

Private Sub StaticButton_Click()
    Me.LowerTextBox.Enabled = Me.StaticButton.Value
    Me.UpperTextBox.Enabled = Me.StaticButton.Value
    
    Me.ComboBox9.Enabled = Not (Me.StaticButton.Value)
    Me.ComboBox10.Enabled = Not (Me.StaticButton.Value)

End Sub

Private Sub VariableButton_Click()
    Me.LowerTextBox.Enabled = Not (Me.VariableButton.Value)
    Me.UpperTextBox.Enabled = Not (Me.VariableButton.Value)
    
    Me.ComboBox9.Enabled = Me.VariableButton.Value
    Me.ComboBox10.Enabled = Me.VariableButton.Value
End Sub
