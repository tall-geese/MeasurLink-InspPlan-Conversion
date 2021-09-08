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

'****************************************************
'******************   PartLib   *********************
'****************************************************

Public Sub ExportQIF(ByRef control As IRibbonControl)
    On Error GoTo 20
    Dim featureArr() As Variant
    Dim routineArr() As String
    Dim partArr() As String
    Dim rev As String
    
    routineArr = Worksheets("PartLib Table").GetRoutineListing()
    If (Not routineArr) = -1 Then
        MsgBox "No Routines Exist"
        Exit Sub
    End If
    
    partArr = Worksheets("Variables").GetPartNumberOrNumbers()
    If (Not partArr) = -1 Then
        MsgBox "No Part Numbers Entered"
        Exit Sub
    End If
    
'    partNum = Worksheets("START HERE").Range("C8").Value 'TODO: we need the logic for multiple part numbers
    rev = Worksheets("START HERE").Range("C10").Value
    If rev = "" Then
        MsgBox "Nothing was set in the Revision field in the START HERE page"
        Exit Sub
    End If
    
    For i = 0 To UBound(routineArr)
        For j = 0 To UBound(partArr)
            'Setting the part number in the START HERE page, deliberately not turning off events
            'The reason is becuase some features will be conditionally hidden when we have certain part Numbers set
            'When they are hidden, CollectFeaturesForExport should pass over them
            Worksheets("START HERE").Range("C8").Value = partArr(j)
            
            featureArr = Worksheets("PartLib Table").CollectFeaturesForExport(routineArr(i))
            If (Not featureArr) = -1 Then
                MsgBox ("Didnt find any characteristics for " & vbCrLf & routineArr(i) & vbCrLf & "No Output")
                GoTo cont
            End If
            
            'TODO: change to include in teh routine name
            Call XMLCreation.CreateXML(featureArr, partArr(j), rev, routineArr(i))
cont:
        Next j
    Next i
    
    
    
    'TODO: put in the logic for iteration through our part numbers in the variables tab or
    'allowing the user to enter in either a range or select the applicable part numbers
    
    'TODO: need the logic for grabbing the desired routines as well
    'User should be able to choose one or all of them
    
    'When we creating a routine, we should be iterating through each feature
    'if if they have Anything in the cell that intersects our routine, then it belongs in there
    
      
20
End Sub



'****************************************************
'********************   Data   **********************
'****************************************************



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


'****************************************************
'***************   Features   ***********************
'****************************************************

'******************   Pivot Features Btn  ***********************

Public Sub PivotFeature(ByRef control As IRibbonControl)
    Set partWS = Worksheets("PartLib Table")
    
    If ActiveSheet.Name = "PartLib Table" Then
        If Not partWS.IsInImmutableRange(ActiveCell) Then
            Dim charCell As Range
            Set charCell = ActiveCell.offset(0, partWS.GetCol("Characteristic Name") - ActiveCell.column)
            If charCell.Value <> vbNullString Then
                Dim headerRange As Range
                Set headerRange = partWS.Range("$D$3:" & partWS.Range("D3").End(xlToRight).Address)
                
                Dim pvtWS As Worksheet
                Set pvtWS = Worksheets("PivotFeature")
                Dim toRange As Range
                Set toRange = pvtWS.Range("B4")
                
                Application.ScreenUpdating = False
                
                pvtWS.Unprotect Password:="3063bqa"
                Call Worksheets("PivotFeature").Cleanup
                
                'Hide/Show our grouped rows, if they are hidden or not on the PartLib Table and vice-versa
                If Not (partWS.Columns(4).EntireColumn.Hidden) Then
                    pvtWS.Rows(3).EntireRow.ShowDetail = True
                End If
                If Not (partWS.Columns(8).EntireColumn.Hidden) Then
                    pvtWS.Rows(8).EntireRow.ShowDetail = True
                End If
                If Not (partWS.Columns(18).EntireColumn.Hidden) Then
                    pvtWS.Rows(18).EntireRow.ShowDetail = True
                End If
                
                
                Dim headerCell As Range
                For Each headerCell In headerRange
                    If toRange.Value = "" Then
                        toRange.Value = headerCell.Value
                        toRange.Interior.color = headerCell.Interior.color
                        toRange.offset(0, -1).Value = "QE"
                        toRange.offset(0, -1).Interior.color = headerCell.Interior.color
                        
                        ThisWorkbook.SetBorders target:=toRange
                        ThisWorkbook.SetBorders target:=toRange.offset(0, -1)
                    End If
                    Set toRange = toRange.offset(1, 0)
                Next headerCell
                
                Dim sample As Integer
                sample = partWS.Range("D3").End(xlToRight).column
                

                Set charCell = Worksheets("PartLib Table").GetCharOrFamily(charCell)
                Worksheets("PivotFeature").SetCharacteristic feature:=charCell, lastCol:=sample
                
                pvtWS.Visible = xlSheetVisible
                pvtWS.Activate
                
                pvtWS.Protect Password:="3063bqa"
            End If
        
        End If
        
    ElseIf ActiveSheet.Name = "PivotFeature" Then
        Worksheets("PartLib Table").Activate
    End If
10
    Application.ScreenUpdating = True
    
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
            For i = 9 To 11
                For Each colCell In varColumns
                    ConditionalFeature.ToleranceFrame.Controls("ComboBox" & i).AddItem (colCell)
                Next colCell
            Next i
            
            ConditionalFeature.Show
        End If
    End If


End Sub

'******************   Hide Features Conditionally Btn  ***********************

Public Sub HideFeaturesCondForm(ByRef control As IRibbonControl)
    Set partWS = Worksheets("PartLib Table")
    
    'set a mfg tolerance for the feature in the given row
    If ActiveSheet.Name = partWS.Name Then
        If TypeName(Selection) = "Range" Then
            Dim label As String
            Dim featureCol As Collection
            Set featureCol = New Collection
            Dim subCell As Range
            For Each subCell In Selection
                If Not partWS.IsInImmutableRange(subCell) Then
                    Dim featureCell As Range
                    'We're going to index from the Characteristic Cell
                    Set featureCell = subCell.offset(0, partWS.GetCol("Characteristic Name") - subCell.column)
                    'Ignore cells w/o Characteristic Names
                    If featureCell.Value = "" Then GoTo cont
                    
                    'If the user did a horizontal collection, we only want to set one feature ONCE
                    If Not ThisWorkbook.IsInColl(featureCol, featureCell) Then
                        featureCol.Add featureCell
                    End If
                End If
cont:
            Next subCell
            
            If featureCol.Count = 0 Then Exit Sub
            If featureCol.Count = 1 Then
                label = featureCol.Item(1).Value
            Else
                label = "*Multiple*"
            End If
            
            Load HideFeatureCond
            HideFeatureCond.FeatureLabel.Caption = label
            
            'Store the address of each applicable cell in the userform
            Dim feature As Range
            For Each feature In featureCol
                HideFeatureCond.Tag = HideFeatureCond.Tag & feature.Address & ","
            Next feature
            HideFeatureCond.Tag = Mid(HideFeatureCond.Tag, 1, Len(HideFeatureCond.Tag) - 1) 'erase the last comma
            
            Dim varColumns As Range
            'Set our ComboBox values with the list of the Variable types
            Set varColumns = Worksheets("Variables").GetVariableColumns()
            For Each colCell In varColumns
                HideFeatureCond.Controls("VariableComboBox").AddItem (colCell)
            Next colCell
            
            HideFeatureCond.Show
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


'****************************************************
'***************   Routines   ***********************
'****************************************************

'******************   Insert Machining Operation Btn  ***********************
Public Sub InsertOperation(ByRef control As IRibbonControl)
    Worksheets("PartLib Table").Activate
    Load CreateRoutinesForm

    CreateRoutinesForm.Show

End Sub


Public Sub DeleteRoutines(ByRef control As IRibbonControl)
    Dim routineArr() As String
    routineArr = Worksheets("PartLib Table").GetRoutineListing()
    If (Not routineArr) = -1 Then 'If there are no routines yet, exit sub
        Exit Sub
    End If
    
    Load DeleteRoutineForm
    For i = 0 To UBound(routineArr)
        DeleteRoutineForm.RoutineComboBox.AddItem routineArr(i)
    Next i
    DeleteRoutineForm.Show
End Sub



'****************************************************
'**************   Dev Tools   ***********************
'****************************************************

Public Sub DisableEvents_Toggle(ByRef control As Office.IRibbonControl, ByRef isPressed As Boolean)
    Application.EnableEvents = Not (isPressed)
End Sub





