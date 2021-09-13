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
                GoTo Cont
            End If
            
            'TODO: change to include in teh routine name
            Call XMLCreation.CreateXML(featureArr, partArr(j), rev, routineArr(i))
Cont:
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


Public Sub ImportRoutineMap(ByRef control As IRibbonControl)
    'TODO: the workbook should be adding its version code on startup
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim extPath As String
    Dim cust As String
    Dim wbPath As String
    Dim readWB As Workbook
    
    
    
    cust = Worksheets("START HERE").Range("C2").Value
    If cust <> "" And fso.FolderExists(DataSources.REPORTS_PATH & cust) Then
        extPath = DataSources.REPORTS_PATH & cust
    Else
        extPath = DataSources.REPORTS_PATH
    End If
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = extPath
        .Title = "Select Routine Map to Import"
        .Show
        
        If .SelectedItems.Count = 0 Then Exit Sub
        wbPath = .SelectedItems.Item(1)
    End With
    
    
    On Error GoTo fileErr
    If (InStr(wbPath, ".xlsx") > 0) Then
        Err.Raise Number:=vbObjectError + 1100, Description:="This should be a new style Routine Map, Version 1.X.X Min"
    End If
    If Not (InStr(wbPath, ".xlsm") > 0) Then
        Err.Raise Number:=vbObjectError + 1100, Description:=""
    End If
    
    
    'Will opening the workbook cause some events to fire that we dont want?
    'Disable events before opening?
    
    Application.EnableEvents = False
    On Error GoTo wbErr
    Set readWB = Workbooks.Open(Filename:=wbPath, UpdateLinks:=False, ReadOnly:=True)
    Application.EnableEvents = True
    
    Dim verCode As String
    On Error GoTo subErr
    With readWB.VBProject.VBComponents("DataSources").CodeModule
        On Error GoTo verErr
        verCode = Split(Split(.Lines(1, .CountOfDeclarationLines), "Const VERSION = " & Chr(34))(1), Chr(34))(0)
    End With
    
    Dim verNums() As String
    verNums = Split(verCode, ".")
    If CDbl(verNums(0) & "." & verNums(1)) < 1.1 Then
        MsgBox (verNums(0) & "." & verNums(1))
        GoTo verErr 'Functionality supported in 1.1.0
    End If
    
    Dim featuresArr() As String
    featuresArr = readWB.GetFeaturesForImport()
    If (Not featuresArr) = -1 Then
        MsgBox "Didn't find any features to set", vbInformation
        GoTo 10
    End If
    
    On Error GoTo featErr
    ThisWorkbook.Worksheets("PartLib Table").ImportRoutineMap (featuresArr)
10
    On Error Resume Next
    readWB.Close SaveChanges:=False
       
    Exit Sub
    
    
fileErr:
    MsgBox "You selected an incorrect file type." & vbCrLf & Err.Description, vbCritical
    Exit Sub
subErr:
    MsgBox "This RoutineMap does not support Import Functionality", vbInformation
    GoTo 10
verErr:
    MsgBox "This Version of the RoutineMap does not support Importing", vbInformation
    GoTo 10
featErr:
    MsgBox "Error encountered when setting feature information", vbInformation
    GoTo 10
wbErr:
    MsgBox "Coudn't Open the Workbook", vbCritical
    On Error Resume Next
    readWB.Close SaveChanges:=False
    Application.EnableEvents = True
    Exit Sub

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
    
    If ActiveSheet.Name = "PartLib Table" Then
        Call Worksheets("PartLib Table").PivotOnFeature
    ElseIf ActiveSheet.Name = "PivotFeature" Then
        Worksheets("PartLib Table").Activate
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
                    If featureCell.Value = "" Then GoTo Cont
                    
                    'If the user did a horizontal collection, we only want to set one feature ONCE
                    If Not ThisWorkbook.IsInColl(featureCol, featureCell) Then
                        featureCol.Add featureCell
                    End If
                End If
Cont:
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







