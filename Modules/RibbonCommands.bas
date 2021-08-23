Attribute VB_Name = "RibbonCommands"
'*************************************************************
'*************************************************************
'*                  RibbonCommands
'*
'*  Define any callback functions for Custom Ribbon.
'*  Invalidate(Refresh) controls using the instance of the cusRibbon, must be stored at ON_Load
'*************************************************************
'*************************************************************


Dim cusRibbon As IRibbonUI



Public Sub Ribbon_OnLoad(uiRibbon As IRibbonUI)
    Set cusRibbon = uiRibbon
    cusRibbon.ActivateTab "mlTab"
End Sub


'***************   Set Data Validations Btn  *********************

Public Sub LoadDataValidations(ByRef control As IRibbonControl)
    Call Validations.OpenDataValidations
    Call Validations.SetDataValidations
End Sub

'***************   Insert New Validation Value Btn  *********************

Public Sub InsertValidationValue(ByRef control As IRibbonControl)
    If ActiveCell.Value = "" Then Exit Sub
    If ActiveSheet.Name = "PartLib Table" Then
        Dim targetCol As Integer
        targetCol = ActiveCell.column
        
        'We only allow additions from the Comments or InspMethods column currently
        If targetCol <> 13 And targetCol <> 14 Then
            MsgBox ("You may only insert Comments or Inspection Methods")
            Exit Sub
        End If
        Call Validations.OpenDataValidations 'if not open, then open it
        If Not Validations.ValidationValueExists(inputVal:=ActiveCell.Value, targetCol:=ActiveCell.column) Then
            Dim userPass As String
                'TODO: change this up to a userform so we can hide the value displayed
            userPass = InputBox("Input Password for RoutineMapDataValidations", "Validations Password")
            If userPass = "" Then Exit Sub
            
            'open the wb in write mode, save the changes and open again in read mode
            Call Validations.InsertNewValidation(newVal:=ActiveCell.Value, targetCol:=ActiveCell.column, userPass:=userPass)
            Call Validations.CloseDataValidations(saveWB:=True)
            Call Validations.OpenDataValidations
        End If
    End If

End Sub


'******************   Set Mfg Tolerance Btn  ***********************

Public Sub SetMfgTolerance(ByRef control As IRibbonControl)

    Set partWS = Worksheets("PartLib Table")
    
    'set a mfg tolerance for the feature in the given row
    If ActiveSheet.Name = partWS.Name Then
        If Not partWS.IsInImmutableRange(ActiveCell) Then
           Dim inspCell As Range
           Set inspCell = ActiveCell.offset(0, partWS.GetCol("Inspection Method") - ActiveCell.column)
           Call partWS.LoadMfgTol(inspCell, 0, 0)
        End If
    End If
End Sub


'******************   Build Variable Formula Btn  ***********************

Public Sub BuildVariableFeatureForm(ByRef control As IRibbonControl)
    Set partWS = Worksheets("PartLib Table")
    
    'set a mfg tolerance for the feature in the given row
    If ActiveSheet.Name = partWS.Name Then
        If Not partWS.IsInImmutableRange(ActiveCell) Then
        
            Load ConditionalFeature
            
            Set inspCell = ActiveCell.offset(0, partWS.GetCol("Characteristic Name") - ActiveCell.column)
            If inspCell.Value = "" Then Exit Sub
            ConditionalFeature.FeatureLabel.Caption = inspCell.Value
            
            Dim varColumns As Range
            Set varColumns = Worksheets("Variables").GetVariableColumns()
            
            For i = 1 To 8
                For Each colCell In varColumns
                    ConditionalFeature.OutputFrame.Controls("ComboBox" & i).AddItem (colCell)
                Next colCell
            Next i
            For i = 9 To 10
                For Each colCell In varColumns
                    ConditionalFeature.ToleranceFrame.Controls("ComboBox" & i).AddItem (colCell)
                Next colCell
            Next i
            
            ConditionalFeature.Show
        End If
    End If


End Sub


