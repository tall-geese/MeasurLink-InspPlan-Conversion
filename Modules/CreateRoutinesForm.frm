VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateRoutinesForm 
   Caption         =   "Create Routines"
   ClientHeight    =   9480
   ClientLeft      =   -320
   ClientTop       =   -1620
   ClientWidth     =   5320
   OleObjectBlob   =   "CreateRoutinesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateRoutinesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'****************   Helper Routines   ***********************
'************************************************************

Dim swsMllFrame As Frame
Dim receivFrame As Frame
Dim assemFrame As Frame
Dim finFrame As Frame

'---------------Option Buttons-----------------
'   The main event
Private Sub CreateRoutinesButton_Click()
    Dim opTag As String
    Dim routineArr() As Variant
    Dim selectedColor As Long
    Dim machiningOp As String
    selectedColor = RGB(255, 192, 192)

    If Me.MillingOptionButton.Value = True Then
        opTag = "MILL"
        GoTo 10
    ElseIf Me.SwissOptionButton.Value = True Then
        opTag = "SWISS"
10
        routineArr = Me.BuildRoutineArray(Me.SwissMillFrame, routineArr)
    
    ElseIf Me.AssemblyOptionButton.Value = True Then
        opTag = "ASSEM"
        routineArr = Me.BuildRoutineArray(Me.AssemblyFrame, routineArr)
    
    ElseIf Me.FinalOptionButton.Value = True Then
        opTag = "FINAL"
        routineArr = Me.BuildRoutineArray(Me.FinalFrame, routineArr)
    
    ElseIf Me.ReceivingOptionButton.Value = True Then
        opTag = "RECEIVE"
        'This is the only one that is different, it needs to be built
        If Trim(Me.OPTextBox.Value) = vbNullString Then GoTo 20
        If Not IsNumeric(Me.OPTextBox.Value) Then GoTo 20
        
        ReDim Preserve routineArr(0)
        routineArr(0) = Replace(Me.FI_OPXX_RECINSPCheckbox.Caption, "XXX", Me.OPTextBox.Value)
    Else
        'Shouldnt be possible
        Exit Sub
    
    End If
    
    'Grab the color if one is selected, if not we always have the default
    For Each colorCtrl In Me.PaletteFrame.Controls
        If colorCtrl.Locked = True Then
            selectedColor = colorCtrl.BackColor
        End If
    Next colorCtrl
    
    'TODO: dont forget to grab the machining operation as well
    machiningOp = Me.OperationTextBox.Value
    Unload Me
    Call Worksheets("PartLib Table").SetRoutines(routineArr, selectedColor, opTag, machiningOp)
    
20
   Unload Me
End Sub

Public Function BuildRoutineArray(ctrFrame As Frame, varArr() As Variant) As Variant()
    For Each ctrl In ctrFrame.Controls
        If ctrl.Value = True Then
            If (Not varArr) = -1 Then
                ReDim Preserve varArr(0)
                varArr(0) = ctrl.Caption
            Else
                ReDim Preserve varArr(UBound(varArr) + 1)
                varArr(UBound(varArr)) = ctrl.Caption
            End If
        End If
    Next ctrl
    
    BuildRoutineArray = varArr


End Function



Public Sub DeactivateOthers(ctrlFrame As Frame)
    For Each ctrl In ctrlFrame.Controls
        ctrl.Enabled = False
    Next ctrl
End Sub
Public Sub ActivateMe(ctrlFrame As Frame)
    For Each ctrl In ctrlFrame.Controls
        ctrl.Enabled = True
    Next ctrl
'    For Each ctrl In ctrlFrame.Controls
'        If ctrl.Name <> "FA_FIRSTCheckbox" And ctrl.Name <> "FA_VISCheckbox" And _
'            ctrl.Name <> "FI_OPXX_RECINSPCheckbox" And ctrl.Name <> "FI_DIMCheckbox" Then
'            ctrl.Enabled = True
'        End If
'    Next ctrl

End Sub


Private Sub OperationTextBox_Change()

End Sub

Private Sub SpinButton1_SpinDown()
    If CInt(Me.OperationTextBox) > 1 Then
        Me.OperationTextBox = CInt(Me.OperationTextBox) - 1
    End If
End Sub

Private Sub SpinButton1_SpinUp()
    If CInt(Me.OperationTextBox) < 4 Then
        Me.OperationTextBox = CInt(Me.OperationTextBox) + 1
    End If
End Sub


Private Sub UserForm_Activate()
    Call SwissOptionButton_Click
End Sub

Private Sub UserForm_Initialize()
    Set swsMllFrame = Me.SwissMillFrame
    Set receivFrame = Me.ReceivingFrame
    Set assemFrame = Me.AssemblyFrame
    Set finFrame = Me.FinalFrame
    
    
End Sub



'---------------Option Buttons-----------------
Private Sub SwissOptionButton_Click()
    Call Me.ActivateMe(swsMllFrame)
    
    Call Me.DeactivateOthers(receivFrame)
    Call Me.DeactivateOthers(assemFrame)
    Call Me.DeactivateOthers(finFrame)
End Sub

Private Sub MillingOptionButton_Click()

    Call Me.ActivateMe(swsMllFrame)
    
    Call Me.DeactivateOthers(receivFrame)
    Call Me.DeactivateOthers(assemFrame)
    Call Me.DeactivateOthers(finFrame)
End Sub
Private Sub ReceivingOptionButton_Click()
    Call Me.ActivateMe(receivFrame)
    
    Call Me.DeactivateOthers(swsMllFrame)
    Call Me.DeactivateOthers(assemFrame)
    Call Me.DeactivateOthers(finFrame)

End Sub
Private Sub AssemblyOptionButton_Click()
    Call Me.ActivateMe(assemFrame)
    
    Call Me.DeactivateOthers(receivFrame)
    Call Me.DeactivateOthers(swsMllFrame)
    Call Me.DeactivateOthers(finFrame)

End Sub
Private Sub FinalOptionButton_Click()
    Call Me.ActivateMe(finFrame)
    
    Call Me.DeactivateOthers(receivFrame)
    Call Me.DeactivateOthers(assemFrame)
    Call Me.DeactivateOthers(swsMllFrame)

End Sub

















'-------------------Buttons-----------------
    'Don't worry about em, they work
Private Sub CommandButton1_Click()
    Me.CommandButton1.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton1.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton2_Click()
    Me.CommandButton2.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton2.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton3_Click()
    Me.CommandButton3.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton3.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton4_Click()
    Me.CommandButton4.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton4.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton5_Click()
    Me.CommandButton5.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton5.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton6_Click()
    Me.CommandButton6.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton6.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton7_Click()
    Me.CommandButton7.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton7.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton8_Click()
    Me.CommandButton8.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton8.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton9_Click()
    Me.CommandButton9.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton9.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton10_Click()
    Me.CommandButton10.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton10.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton11_Click()
    Me.CommandButton11.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton11.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton12_Click()
    Me.CommandButton12.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton12.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton13_Click()
    Me.CommandButton13.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton13.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton14_Click()
    Me.CommandButton14.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton14.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton15_Click()
    Me.CommandButton15.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton15.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton16_Click()
    Me.CommandButton16.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton16.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton17_Click()
    Me.CommandButton17.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton17.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton18_Click()
    Me.CommandButton18.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton18.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton19_Click()
    Me.CommandButton19.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton19.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton20_Click()
    Me.CommandButton20.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton20.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub
Private Sub CommandButton21_Click()
    Me.CommandButton21.Locked = True
    For Each ctrl In Me.PaletteFrame.Controls
        If ctrl.name <> Me.CommandButton21.name Then
            ctrl.Locked = Falsed
        End If
    Next ctrl
End Sub


