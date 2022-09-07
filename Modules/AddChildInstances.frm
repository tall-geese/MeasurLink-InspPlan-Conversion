VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddChildInstances 
   Caption         =   "Add Child Instances"
   ClientHeight    =   3465
   ClientLeft      =   -240
   ClientTop       =   -1050
   ClientWidth     =   4770
   OleObjectBlob   =   "AddChildInstances.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddChildInstances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lowerBoundInst As Integer


Private Sub NumInstancesSpin_SpinDown()
    If CInt(Me.NumInstancesBox.Value) = lowerBoundInst Then Exit Sub
    
    Me.NumInstancesBox.Value = CInt(Me.NumInstancesBox.Value) - 1
    If CInt(Me.NumInstancesBox.Value) = lowerBoundInst Then
        Me.SubmitBtn.Enabled = False
    End If
    
End Sub

Private Sub NumInstancesSpin_SpinUp()
    Me.NumInstancesBox.Value = CInt(Me.NumInstancesBox.Value) + 1
    If CInt(Me.NumInstancesBox.Value) > lowerBoundInst Then
        Me.SubmitBtn.Enabled = True
    End If
End Sub

Private Sub SubmitBtn_Click()
    If Me.SubmitBtn.Enabled = True And CInt(Me.NumInstancesBox.Value) > lowerBoundInst Then
        Me.Tag = Me.NumInstancesBox.Value
    End If
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    Me.NumInstancesBox.Value = lowerBoundInst
End Sub




