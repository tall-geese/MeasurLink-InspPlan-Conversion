VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MapRoutinesForm 
   Caption         =   "Auto Map Routines"
   ClientHeight    =   4860
   ClientLeft      =   -585
   ClientTop       =   -2370
   ClientWidth     =   5295
   OleObjectBlob   =   "MapRoutinesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MapRoutinesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MapRoutinesBtn_Click()
    Me.Tag = "True"
    Me.Hide
End Sub
