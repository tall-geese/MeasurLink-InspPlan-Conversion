VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteRoutineForm 
   Caption         =   "Delete Routine"
   ClientHeight    =   1896
   ClientLeft      =   -180
   ClientTop       =   -768
   ClientWidth     =   3072
   OleObjectBlob   =   "DeleteRoutineForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteRoutineForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Sub LoadRoutines(routineArr() As String)
'    For i = 0 To UBound(routineArr)
'        Me.RoutineComboBox.AddItem routineArr(i)
'    Next i
'End Sub



'The Main Event Here
Private Sub DeleteRoutineButton_Click()
    If Me.RoutineComboBox.Value = "" Then GoTo 10
    Call Worksheets("PartLib Table").DeleteRoutine(Me.RoutineComboBox.Value)
    Unload Me
10
    
End Sub

Private Sub UserForm_Click()

End Sub
