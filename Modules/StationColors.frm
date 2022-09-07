VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StationColors 
   Caption         =   "Choose Color for the Cell"
   ClientHeight    =   4185
   ClientLeft      =   -15
   ClientTop       =   -150
   ClientWidth     =   6705
   OleObjectBlob   =   "StationColors.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StationColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cleared_fields As Boolean
Public colorResult As String

'*************************************************************
'*************************************************************
'*                  Extra Callbacks
'*************************************************************
'*************************************************************


Private Sub SubmitColorButton_Click()
    'Check the field sand make sure that their values make sense
    If Me.Field_Red.Value <> vbNullString And Me.Field_Green.Value <> vbNullString And Me.Field_Blue.Value <> vbNullString Then
        colorResult = "rgb(" & Me.Field_Red.Value & "," & Me.Field_Green.Value & "," & Me.Field_Blue.Value & ")"
        GoTo finish_unload
    Else
        For Each contr In Me.Controls
            If TypeName(contr) = "CommandButton" Then
                If contr.Locked = True Then
                    colorResult = contr.Tag
                    GoTo finish_unload
                End If
            End If
        Next contr
    End If
    
    'TODO: have some kind of error message here, not supposed to get here
    MsgBox "You need to Enter color information or select a Color", vbInformation
    Exit Sub
    
finish_unload:
    'Need to make sure that we have a certain exit code status
    Unload Me
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 0 Then
        'Pass the value back to the calling StationModify form
        StationModify.Color_Received RGB:=colorResult
    End If
End Sub

Public Sub ColorSeed(red As String, green As String, blue As String)
    Me.Field_Red.Value = red
    Me.Field_Green = green
    Me.Field_Blue = blue
End Sub





'*************************************************************
'*************************************************************
'*                  RGB Field CallBacks
'*************************************************************
'*************************************************************

Private Sub Unlock_Fields()
    cleared_fields = True
    
    Me.Field_Blue.Value = ""
    Me.Field_Green.Value = ""
    Me.Field_Red.Value = ""
End Sub

Private Sub Validate_color(contr As MSForms.control)
    Dim val As Integer
    On Error Resume Next
    
    val = CInt(contr.Value)
    If Err.Number <> 0 Then
        contr.Value = ""
    End If
    
    If Not ((val >= 0) And (val <= 255)) Then
        contr.Value = ""
    End If
End Sub


Private Sub Field_Blue_Enter()
    Call Unlock_buttons
End Sub

Private Sub Field_Blue_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Validate_color Me.Field_Blue
End Sub

Private Sub Field_Blue_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not cleared_fields Then Call Unlock_Fields
End Sub


Private Sub Field_Green_Enter()
    Call Unlock_buttons
End Sub

Private Sub Field_Green_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Validate_color Me.Field_Green
End Sub

Private Sub Field_Green_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not cleared_fields Then Call Unlock_Fields
End Sub

Private Sub Field_Red_Enter()
    Call Unlock_buttons
End Sub

Private Sub Field_Red_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Validate_color Me.Field_Red
End Sub

Private Sub Field_Red_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not cleared_fields Then Call Unlock_Fields
End Sub


'*************************************************************
'*************************************************************
'*                  Color Button CallBacks
'*************************************************************
'*************************************************************
Private Sub Set_Lock(contr As MSForms.control)

    For Each sub_contr In Me.Controls
        If TypeName(sub_contr) = "CommandButton" Then
            If Not ((sub_contr.name = contr.name) Or (sub_contr.name = "SubmitColorButton")) Then
                sub_contr.Locked = False
            End If
        End If
    Next sub_contr
    
    contr.Locked = True
    
    Me.Field_Red.Value = ""
    Me.Field_Green.Value = ""
    Me.Field_Blue.Value = ""
    
    'Lock all other buttons except the one above and the submit button
    'and I guess we should clear the Color fields if they have values in them

End Sub

Private Sub Unlock_buttons()
    'TODO: Should be called by the fields to unlock the buttons to be chosen once again...
    For Each contr In Me.Controls
        If TypeName(contr) = "CommandButton" Then
            contr.Locked = False
        End If
    Next contr
    
End Sub


Private Sub Color1_Click()
    Set_Lock contr:=Me.Color1
End Sub

Private Sub Color2_Click()
    Set_Lock contr:=Me.Color2
End Sub

Private Sub Color3_Click()
    Set_Lock contr:=Me.Color3
End Sub

Private Sub Color4_Click()
    Set_Lock contr:=Me.Color4
End Sub

Private Sub Color5_Click()
    Set_Lock contr:=Me.Color5
End Sub

Private Sub Color6_Click()
    Set_Lock contr:=Me.Color6
End Sub

Private Sub Color7_Click()
    Set_Lock contr:=Me.Color7
End Sub

Private Sub Color8_Click()
    Set_Lock contr:=Me.Color8
End Sub

Private Sub Color9_Click()
    Set_Lock contr:=Me.Color9
End Sub

Private Sub Color10_Click()
    Set_Lock contr:=Me.Color10
End Sub

Private Sub Color11_Click()
    Set_Lock contr:=Me.Color11
End Sub

Private Sub Color12_Click()
    Set_Lock contr:=Me.Color12
End Sub

Private Sub Color13_Click()
    Set_Lock contr:=Me.Color13
End Sub

Private Sub Color14_Click()
    Set_Lock contr:=Me.Color14
End Sub


Private Sub UserForm_Click()

End Sub

