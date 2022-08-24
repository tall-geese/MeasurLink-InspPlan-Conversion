VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} INSERTview 
   Caption         =   "Information to INSERT"
   ClientHeight    =   1980
   ClientLeft      =   -105
   ClientTop       =   -390
   ClientWidth     =   930
   OleObjectBlob   =   "INSERTview.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "INSERTview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 0 Then
        RibbonCommands.add_custom_fields_valid = True
    End If
End Sub
