VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateMfgTolerance 
   Caption         =   "Create Mfg Tolernace"
   ClientHeight    =   3795
   ClientLeft      =   -420
   ClientTop       =   -1760
   ClientWidth     =   5440
   OleObjectBlob   =   "CreateMfgTolerance.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateMfgTolerance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub NoteTextBox_Change()

End Sub

Private Sub SubmitButton_Click()
    Dim inputs(2) As Variant
    inputs(0) = Me.LTolTextBox.Text
    inputs(1) = Me.TargetTextBox.Text
    inputs(2) = Me.UTolTextBox.Text
    
    Dim charAddress As String
    charAddress = Me.Tag
    
    For i = 0 To 2
        If inputs(i) = "" Then
            MsgBox ("A value was missing, Can't set any values")
            Exit Sub
        End If
    Next i
    
    Dim optMsg As Variant
    If Me.NoteTextBox <> vbNullString Then optMsg = Me.NoteTextBox
    
    Unload Me
    
    Call ThisWorkbook.Worksheets("PartLib Table").SetMfgTol(charAddress:=charAddress, inputs:=inputs, tolNote:=optMsg)
    
End Sub

Private Sub UserForm_Initialize()
'    Debug.Print (testCell.Value & " we're inside the userform!!")
End Sub
