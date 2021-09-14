VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "btnClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public usrForm As Object

Private Sub btn_Click()
    usrForm.Hide
End Sub

