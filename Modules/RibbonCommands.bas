Attribute VB_Name = "RibbonCommands"
'*************************************************************
'*************************************************************
'*                  RibbonCommands
'*
'*  Define any callback functions for Custom Ribbon.
'*  Invalidate(Refresh) controls using the instance of the cusRibbon, must be stored at ON_Load
'*************************************************************
'*************************************************************


Private cusRibbon As IRibbonUI

Private ribbonMsg As String

Private partCombo_Enabled As Boolean
Private partCombo_PartList() As String
Private partCombo_TextField As String

Private json_parts_info As Object 'JSON object returned from API
Private json_part_lib As Object  'JSON object made by PartLib.Build_JSON_of_Parts

Private toggle_viewCustomFields As Boolean

Public add_custom_fields_valid As Boolean



Public Sub Ribbon_OnLoad(uiRibbon As IRibbonUI)
    Set cusRibbon = uiRibbon
    cusRibbon.ActivateTab "mlTab"
End Sub

Private Sub InvalidateControl(controlName As String)
    On Error Resume Next
    cusRibbon.InvalidateControl controlName
    If cusRibbon Is Nothing Then
        MsgBox "Reference to the Excel Ribbon has been lost and Custom Ribbon Controls may not function properly" & vbCrLf & vbCrLf _
            & "This can happen after an Error occurs or certain parts of the code have been edited" & vbCrLf _
            & "Save your work and re-open the workbook and try again", vbExclamation
    End If
End Sub

'****************************************************
'******************   PartLib   *********************
'****************************************************

Public Sub ExportQIF(ByRef control As IRibbonControl)
    On Error GoTo 20
    Dim featureArr() As Variant
    Dim routineArr() As String
    Dim partArr() As String
    Dim routineWarnings() As String
    Dim rev As String
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    
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
    
    rev = Worksheets("START HERE").Range("C10").Value
    If rev = "" Then
        MsgBox "Nothing was set in the Revision field in the START HERE page"
        Exit Sub
    End If
    
        'Create the output folder if one doe not exist
    On Error GoTo IOerror
    If Not fso.FolderExists(ThisWorkbook.path & "\Output") Then
        Dim result As Integer
        result = MsgBox("There is no Output folder" & vbCrLf & "Would you like to create one?", vbYesNo)
        If result <> vbYes Then Exit Sub
        
        fso.CreateFolder (ThisWorkbook.path & "\Output")
    End If

   
        'Create the partNumber folder for QIF files belonging to that partNumber
    For i = 0 To UBound(partArr)
        On Error GoTo IOerror
        Worksheets("START HERE").Range("C8").Value = partArr(i)
            
        If fso.FolderExists(ThisWorkbook.path & "\Output\" & partArr(i)) Then
            fso.DeleteFolder (ThisWorkbook.path & "\Output\" & partArr(i))
        End If
        
        fso.CreateFolder (ThisWorkbook.path & "\Output\" & partArr(i))
    
        For j = 0 To UBound(routineArr)
            'Setting the part number in the START HERE page, deliberately not turning off events
            'The reason is becuase some features will be conditionally hidden when we have certain part Numbers set
            'When they are hidden, CollectFeaturesForExport should pass over them
            Dim k As Integer
            
            featureArr = Worksheets("PartLib Table").CollectFeaturesForExport(routineArr(j))
                'If a routine didnt have any features set for inspection, let the user know, but only once per unique Routine name
            If (Not featureArr) = -1 Then
                If (Not routineWarnings) = -1 Then
                    MsgBox ("Didnt find any characteristics for " & vbCrLf & routineArr(j) & vbCrLf & "No Output")
                    ReDim Preserve routineWarnings(0)
                    routineWarnings(0) = routineArr(j)
                Else
                    Dim rtFound As Boolean
                    For k = 0 To UBound(routineWarnings)
                        If routineWarnings(k) = routineArr(j) Then rtFound = True
                    Next k
                    
                    If rtFound = False Then
                        MsgBox ("Didnt find any characteristics for " & vbCrLf & routineArr(j) & vbCrLf & "No Output")
                        ReDim Preserve routineWarnings(UBound(routineWarnings) + 1)
                        routineWarnings(UBound(routineWarnings)) = routineArr(j)
                    End If
                    
                End If
                GoTo Cont
            End If
            
            
            'stop here to check, if we have a FI_DIM routine or we have a RECINSP operation then
            'if we have only attribute features, then we should error and let the user know, but still can continue....
            
            If InStr(routineArr(j), "FI_DIM") > 0 Or InStr(routineArr(j), "RECINSP") > 0 Then
                Dim hasVariable As Boolean
                For k = 1 To UBound(featureArr, 2)
                    If featureArr(1, k) = "Variable" Then hasVariable = True
                Next k
                
                If Not hasVariable Then
                    MsgBox "Routine: " & routineArr(j) & vbCrLf & "has no variable features for inspection" & vbCrLf & vbCrLf _
                            & "Double check, as this is most likely incorrect", vbInformation
                End If
                
                hasVariable = False
            End If
            
            On Error GoTo XMLerror
            
            Call XMLCreation.CreateXML(featureArr, partArr(i), rev, routineArr(j))
Cont:
        Next j
    Next i
    
    
      
20
    Exit Sub
    
    
IOerror:
    MsgBox "Couldn't Create/Delete directories" & vbCrLf & "You may not have the proper read/write permissions" & vbCrLf _
                & "Or the part number may contain an illegal character for Windows", vbCritical
                
    Exit Sub
    
XMLerror:
    Exit Sub

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
'    Debug.Print (CDbl(verNums(0) & "." & verNums(1)))
    If CDbl(verNums(0) & "." & verNums(1)) < 1.1 Then
        MsgBox (verNums(0) & "." & verNums(1))
        GoTo verErr 'Initial Import Functionality supported in 1.1.0
    End If
    
    Dim featuresArr() As String
    featuresArr = readWB.GetFeaturesForImport()
    If (Not featuresArr) = -1 Then
        MsgBox "Didn't find any features to set", vbInformation
        GoTo 10
    End If
    
    On Error GoTo featErr
    
    Call ThisWorkbook.Worksheets("PartLib Table").Outline.ShowLevels(RowLevels:=2)
    Call ThisWorkbook.Worksheets("PartLib Table").ImportRoutineMap(featuresArr, readWB.name, readWB.path)
10
    On Error Resume Next
    readWB.Close SaveChanges:=False
       
    Exit Sub
    
    
fileErr:
    MsgBox "You selected an incorrect file type." & vbCrLf & Err.Description, vbCritical
    Exit Sub
subErr:
    MsgBox "This RoutineMap either does not support Import Functionality" & vbCrLf & "Or you may not have 'Trust Access to VBA Project Model'" _
                & " Enabled in your Excel Settings", vbInformation
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


'***************   Import New Data Validations Btn  *********************

Public Sub ImportDataValidations(ByRef control As IRibbonControl)
    Call ImportValidationValues
End Sub


        '*********   Deprecated  ***************
'***************   Set Data Validations Btn  *********************

'Public Sub LoadDataValidations(ByRef control As IRibbonControl)
'    Call Validations.OpenDataValidations
'    Call Validations.SetDataValidations
'End Sub



'***************   Insert New Validation Value Btn  *********************

Public Sub InsertValidationValue(ByRef control As IRibbonControl)
    If ActiveCell.Value = "" Then Exit Sub
    If ActiveSheet.name = "PartLib Table" Then
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
            Call ImportValidationValues
        End If
    End If

End Sub


'****************************************************
'***************   Features   ***********************
'****************************************************

'******************   Pivot Features Btn  ***********************

Public Sub PivotFeature(ByRef control As IRibbonControl)
    
    If ActiveSheet.name = "PartLib Table" Then
        Call Worksheets("PartLib Table").PivotOnFeature
    ElseIf ActiveSheet.name = "PivotFeature" Then
        Worksheets("PartLib Table").Activate
    End If
End Sub

'******************   Add Child Instances Btn  ***********************

Public Sub AddChildFeatures(ByRef control As IRibbonControl)
    
    If ActiveSheet.name <> "PartLib Table" Then Exit Sub
    Call Worksheets("PartLib Table").AddFeatureInstances(ActiveCell)
    
End Sub


'******************   Build Variable Formula Btn  ***********************

Public Sub BuildVariableFeatureForm(ByRef control As IRibbonControl)
    Set partWS = Worksheets("PartLib Table")
    
    'set a mfg tolerance for the feature in the given row
    If ActiveSheet.name = partWS.name Then
        If Not partWS.IsInImmutableRange(ActiveCell) Then
        
            Load ConditionalFeature
            
            Set inspCell = ActiveCell.Offset(0, partWS.GetCol("Characteristic Name") - ActiveCell.column)
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
    If ActiveSheet.name = partWS.name Then
        If TypeName(Selection) = "Range" Then
            Dim Label As String
            Dim featureCol As Collection
            Set featureCol = New Collection
            Dim subCell As Range
            For Each subCell In Selection
                If Not partWS.IsInImmutableRange(subCell) Then
                    Dim featureCell As Range
                    'We're going to index from the Characteristic Cell
                    Set featureCell = subCell.Offset(0, partWS.GetCol("Characteristic Name") - subCell.column)
                    'Ignore cells w/o Characteristic Names
                    If featureCell.Value = "" And featureCell.formula = "" Then GoTo Cont
                    
                    'If the user did a horizontal collection, we only want to set one feature ONCE
                    If Not ThisWorkbook.IsInColl(featureCol, featureCell) Then
                        featureCol.Add featureCell
                    End If
                End If
Cont:
            Next subCell
            
            If featureCol.Count = 0 Then Exit Sub
            If featureCol.Count = 1 Then
                Label = featureCol.Item(1).Value
            Else
                Label = "*Multiple*"
            End If
            
            
                'Validate no hiddens in the Collection
            Dim hiddenAlready As Boolean
            Dim hiddenIndexes As Collection
            Set hiddenIndexes = New Collection
            Dim i As Integer
            i = 1
            For Each feature In featureCol
                    'If a characteristic name has a formula in it
                If InStr(feature.formula, "=IF(") > 0 Then
                    If IsNumeric(Right(feature.Value, 2)) Then
                        If CInt(Right(feature.Value, 2)) <= 1 Then   'And its not becuase it is a child feature...
                            hiddenAlready = True
                            hiddenIndexes.Add i
                        End If
                    Else
                        hiddenAlready = True
                        hiddenIndexes.Add i
                    End If
                End If
                i = i + 1
            Next feature
            
                'If the selected features are hidden features
            If hiddenAlready Then
                result = MsgBox("Some of the Features Selected appear to be conditionally hidden" & vbCrLf _
                                & "Would you like to attempt to remove conditional hiding from these features?", vbYesNo)
                If result = vbNo Then Exit Sub
                
                If result = vbYes Then
                    On Error GoTo resetFeatErr
                    Application.EnableEvents = False
                    
                    For Each ind In hiddenIndexes
                        Dim featCell As Range
                        Set featCell = featureCol.Item(ind)
                        Call Worksheets("PartLib Table").UnsetHiding(featCell:=featCell)
                    Next ind
                End If
                
                'If the selected features are not already hidden
            Else
                Load HideFeatureCond
                HideFeatureCond.FeatureLabel.Caption = Label
                
                'Store the address of each applicable cell in the userform
'                Dim feature As Range
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
    End If
10
    Application.EnableEvents = True
    Exit Sub

resetFeatErr:
    MsgBox "Something went wrong with " & featureCol.Item(ind) & vbCrLf & "couldn't strip the value", vbCritical
    GoTo 10
    Exit Sub

End Sub

'******************   Set Mfg Tolerance Btn  ***********************

Public Sub SetMfgTolerance(ByRef control As IRibbonControl)

    Set partWS = Worksheets("PartLib Table")
    
    'set a mfg tolerance for the feature in the given row
    If ActiveSheet.name = partWS.name Then
        If Not partWS.IsInImmutableRange(ActiveCell) Then
           Dim inspCell As Range
           Set inspCell = ActiveCell.Offset(0, partWS.GetCol("Inspection Method") - ActiveCell.column)
           Call partWS.LoadMfgTol(inspCell, 0, 0)
        End If
    End If
End Sub


'******************   Apply Custom Sort Btn  ***********************
Public Sub ApplyCustomSort(ByRef control As IRibbonControl)
    Set partWS = Worksheets("PartLib Table")
    
    If ActiveSheet.name = partWS.name Then
        partWS.Activate
    End If
    
    Call partWS.SortFeatures

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

'******************   Delete Routines Btn  ***********************
Public Sub DeleteRoutines(ByRef control As IRibbonControl)
    Dim colors() As Variant
    Dim routines() As Variant
    Dim selectedRoutines() As String
    
    On Error GoTo rtErr
    routines = Worksheets("PartLib Table").GetRoutinesAndColors(colors)
    
    'TODO: error handle here before trying to delete routines???
    selectedRoutines = ThisWorkbook.BuildRoutineForm(routines:=routines, colors:=colors, frmHeader:="Select Routine(s) to Delete", _
                            btnCaption:="Delete")
    If (Not selectedRoutines) = -1 Then Exit Sub
    
    For i = 0 To UBound(selectedRoutines)
        Call Worksheets("PartLib Table").DeleteRoutine(selectedRoutines(i))
    Next i
    
    Exit Sub
rtErr:
    MsgBox "Couldn't read the Value or Color of a Routine", vbCritical
    Exit Sub
End Sub

'******************   Optimize Inspections Btn  ***********************
Public Sub OptimizeInspections(ByRef control As IRibbonControl)
    Dim charArr() As String
    Dim uniqueOps As Collection
    Set uniqueOps = New Collection
    Dim routines() As Variant
    Dim colors() As Variant
    
    'Collection of routines that we should handle differently
        'Used for the ALL or FAIs Operations
    Dim skipCollection As Collection
    Set skipCollection = New Collection
    skipCollection.Add ("FA_LASR")
    skipCollection.Add ("FI_CMM")
    skipCollection.Add ("FI_RAM")
    skipCollection.Add ("FI_SYLVAC")
    skipCollection.Add ("FI_COMPAR")
    
    Dim FACollection As Collection
    Set FACollection = New Collection
    FACollection.Add ("FA_FIRST")
    FACollection.Add ("FA_VIS")
    FACollection.Add ("FA_MINI")
    
    Dim allRts As Collection
    Set allRts = New Collection
    allRts.Add ("FA_FIRST")
    allRts.Add ("FA_MINI")
    allRts.Add ("IP_BENCH")
    allRts.Add ("FA_VIS")
    allRts.Add ("FI_DIM")
    allRts.Add ("FI_VIS")
    
    
    
        'Assert we have information filled out, and grab the characteristics and operations
    If Worksheets("PartLib Table").IsValidForInspection(charArr, uniqueOps) Then
        routines = Worksheets("PartLib Table").GetRoutinesAndColors(colors)
        
        'Constructing the form...
            'if the amount of unique Operations is 0 or 1, Then we only need to build a single frame of the applicable routines
            'and list the balloon numbers affected above it
        If Not (ThisWorkbook.BuildOptimizeInspectionForm(charArr, uniqueOps, routines, skipCollection, FACollection, allRts)) Then Exit Sub
        'For each opName in orutines
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
            'Cleanup the routines from any previous mapping
        For i = 1 To UBound(routines, 1)
            'SWISS/MILL = routines(i,0)
            'Array of routineNames = routines(i,1)(k)
            For j = 0 To UBound(routines(i, 1))
                Worksheets("PartLib Table").ClearRoutineMapping (routines(i, 1)(j))  'Begin with erasing all of the old mappings
            Next j
        Next i
        For i = 0 To UBound(charArr)
            For j = 1 To UBound(routines, 1)
                If charArr(i, 4) = routines(j, 0) Then  'If character belongs to the assigned operation (like SWISS)
                    For k = 0 To UBound(routines(j, 1))  'Then for each of the routines in that block
                            'Evaluate need for inspection for the given routine and frequency(ies) and method
                        Worksheets("PartLib Table").AssignAsInspection charAddy:=charArr(i, 5), frequency:=charArr(i, 3), _
                                                        routineName:=CStr(routines(j, 1)(k)), inspMethod:=charArr(i, 2), _
                                                        operation:=charArr(i, 4)
                    Next k
                End If
            Next j
        Next i
    End If

End Sub



            '*********   Deprecated  ***************

''******************   Optimize Offsetables Btn  ***********************
'Public Sub OptimizeOffsetables(ByRef control As IRibbonControl)
'    Dim colors() As Variant
'    Dim routines() As Variant
'    Dim selectedRoutines() As String
'    Dim instructions As String
'    Dim offsetExclusions As Collection
'    instructions = "The Selected Routines will have all of the characteristics designated for inspection changed to " & vbCrLf _
'                & "Should Fall In (X). Then the Offsettable features of the smallest tolerance ranges will be set (O)" & vbCrLf _
'                & vbCrLf & "*** Note that Most FA_ and all FI_ routines will always be SFI's"
'    'TODO: come back later and create a collection of routineNames to ignore
'    Dim likeList As Collection
'    Set likeList = New Collection
'    likeList.Add "FA_FIRST"
'    likeList.Add "FA_MINI"
'    likeList.Add "IP_"
'
'
'    On Error GoTo rtErr
'    routines = Worksheets("PartLib Table").GetRoutinesAndColors(colors)
'
'
'
'    Set offsetExclusions = New Collection
'    offsetExclusions.Add "IP_LAST"
'    offsetExclusions.Add "FA_SYLVAC"
'    offsetExclusions.Add "FA_RAMPROG"
'    offsetExclusions.Add "FA_CMM"
'    offsetExclusions.Add "FI_ALL"
'    offsetExclusions.Add "FA_VIS"
'
'    Call Worksheets("PartLib Table").OptimizeRoutineOffsetables(routines, offsetExclusions)
'
'    Exit Sub
'rtErr:
'    MsgBox "Couldn't read the Value or Color of a Routine", vbCritical
'    Exit Sub
'End Sub


'****************************************************
'************   Custom Fields   *********************
'****************************************************

'******************  ViewCustomFields ToggleButton   ***********************

Public Sub viewCustomFields_Toggle(ByRef control As Office.IRibbonControl, ByRef isPressed As Boolean)
    
    toggle_viewCustomFields = isPressed
    
    Dim partNums() As String
    partNums = GetParts_or_SetError()
    
    'If the toggle button was un-pressed or the list of part numbers was not valid
    If Not toggle_viewCustomFields Or (Not partNums) = -1 Then
        ResetViewControls
        Exit Sub
    End If
        
    'Otherwise its a valid list of partNumbers
    ClearRibbonNotification
    partCombo_Enabled = True
    
    
    On Error GoTo Json_Parts_err
    
    Set json_parts_info = HTTPconnections.GetPartsInfo(partNums)
    If json_parts_info Is Nothing Then
        ResetViewControls notification_msg:="     Error:    Parts Don't Exist in MeasurLink ?"
        Exit Sub
    End If
    
    partCombo_PartList = partNums
    partCombo_TextField = partNums(0)

    InvalidateControl "partCombo"
    
    Dim first_part As Object
    With ThisWorkbook.Worksheets("View_CustomFields")
        .Visible = True
        .Activate
        Debug.Print (json_parts_info Is Nothing)
        Set first_part = json_parts_info(GetPartIndex(partNums(0), json_parts_info))
        .LoadPartInformation first_part
    End With
    
    Exit Sub

Json_Parts_err:
    If Err.Number = vbObjectError + 5100 Then   'The first part in the list didnt have anything returned from MeasurLink
        MsgBox "Information on Part#: " & Text & vbCrLf & "Not Found in MeasurLink", vbInformation
    ElseIf Err.Number = vbObjectError + 6000 Then  'Error sending Http
        MsgBox "HTTP Err: viewCustomFields_Toggle" & vbCrLf & Err.Description
    Else
        MsgBox Err.Description
    End If

End Sub


Public Sub viewCustomFields_OnGetPressed(ByRef control As IRibbonControl, ByRef ReturnedValue As Variant)
    ReturnedValue = toggle_viewCustomFields
End Sub

'Called by View_CustomFields when hiding as well as other functions around here
Public Sub ResetViewControls(Optional notification_msg As String)
    'Reset Variables
    Erase partCombo_PartList
    Set json_parts_info = Nothing
    
    'Set optional error message to user
    If notification_msg = vbNullString Then notification_msg = " "
    ribbonMsg = notification_msg
    
    'Reset controls variables
    toggle_viewCustomFields = False
    partCombo_Enabled = False
    partCombo_TextField = vbNullString
    
    'Reload the controls
    InvalidateControl "notificationLabel"
    InvalidateControl "partCombo"
    InvalidateControl "viewCustomFields"
    
    'Hide the CustomFields Sheet if not hidden already
    ThisWorkbook.Worksheets("View_CustomFields").Visible = False
End Sub

Public Sub ClearRibbonNotification()
    ribbonMsg = " "
    InvalidateControl "notificationLabel"
End Sub


'******************   Add Custom Fields Button ***********************
Public Sub AddCustomFields_OnAction(ByRef control As Office.IRibbonControl)
    'Check that the User is even allowed to do this action first
    If Get_API_Key() = vbNullString Then Exit Sub


    'Load up the INSERTform and fetch the values that we want to insert

    Dim json_parts_api As Object, json_parts_map As Object
    Dim partNums() As String
    
    partNums = GetParts_or_SetError()
    If (Not partNums) = -1 Then Exit Sub
    
    Set json_parts_api = HTTPconnections.GetPartsInfo(partNums)
    Set json_parts_map = Worksheets("PartLib Table").Build_JSON_of_Parts(partNums)
    
    On Error GoTo insert_form_Err
    
    Load INSERTform
    
    INSERTform.BuildListArray json_parts_api, json_parts_map
    
    INSERTform.Show
    
    'Can we see what the exit status was? Do they actually want to go and
        'Insert the values or did they just close out of the form??
    
    Exit Sub
insert_form_Err:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

    'Passed back by the INSERTform after it closes
Public Sub ParseArray_ForUpload(json_parts_info As Object)
    Dim output As String, backup As String
    output = JsonConverter.ConvertToJson(json_parts_info, Whitespace:=3)
    backup = output
    
    'Make a copy that is human readable
    output = Replace(output, Chr(34) & "customFieldId" & Chr(34) & ": 13", _
                            Chr(34) & "customField" & Chr(34) & ":Balloon")
    output = Replace(output, Chr(34) & "customFieldId" & Chr(34) & ": 15", _
                            Chr(34) & "customField" & Chr(34) & ":Pins/Gauges")
    output = Replace(output, Chr(34) & "customFieldId" & Chr(34) & ": 3", _
                            Chr(34) & "customField" & Chr(34) & ":Attribute Tolerance")
    output = Replace(output, Chr(34) & "customFieldId" & Chr(34) & ": 8", _
                            Chr(34) & "customField" & Chr(34) & ":Comments")
    output = Replace(output, Chr(34) & "customFieldId" & Chr(34) & ": 11", _
                            Chr(34) & "customField" & Chr(34) & ":Insp. Method")
    output = Replace(output, Chr(34) & "customFieldId" & Chr(34) & ": 12", _
                            Chr(34) & "customField" & Chr(34) & ":Insp. Frequency")
    output = Replace(output, Chr(34) & "customFieldId" & Chr(34) & ": 16", _
                            Chr(34) & "customField" & Chr(34) & ":Char. Description")

    
    'Give the user a last chance to review the Fields they are about to insert.
    Load INSERTview
    INSERTview.json_label.Caption = output
    INSERTview.ScrollHeight = INSERTview.json_label.Height + 90
    
    INSERTview.Show
    
    'If the user said to proceed with adding the fields
    If add_custom_fields_valid Then
        add_custom_fields_valid = False
    Else
        Exit Sub
    End If
    
    Dim api_key As String, resp As String
    api_key = Get_API_Key()
    resp = HTTPconnections.AddCustomFields(payload:=backup, api_key:=api_key)
    
    If resp <> vbNullString Then
        MsgBox "Custom Fields Added Successfully!", vbInformation
    End If
    
End Sub


'******************   PartNumber ComboBox ***********************

Public Sub partCombo_OnChange(ByRef control As Office.IRibbonControl, ByRef Text As Variant)
    
    'Check if this was a valid input (User chose value from the drop-down list)
    If IsError(Application.Match(Text, partCombo_PartList, 0)) Then GoTo reset_controls
    
    On Error GoTo Part_Handle_Err
    Dim part_norev As String
    part_norev = Split(CStr(Text), "_")(0)
    ThisWorkbook.Worksheets("START HERE").SetPartNumber partNum:=part_norev
    
    On Error GoTo Json_Part_err
    ThisWorkbook.Worksheets("View_CustomFields").LoadPartInformation json_parts_info(GetPartIndex(CStr(Text), json_parts_info))
    
    Exit Sub
    
    'If we got here, then we couldnt find the part in the part list, we should reset everything
reset_controls:
    ResetViewControls
    MsgBox "Not A Valid Input", vbInformation
    Exit Sub
    
Json_Part_err:
    If Err.Number = vbObjectError + 5100 Then
        ThisWorkbook.Worksheets("View_CustomFields").Cleanup
        MsgBox "Information on Part#: " & Text & vbCrLf & "Not Found in MeasurLink", vbInformation
    Else
        MsgBox Err.Description, vbCritical
    End If
    
    Exit Sub
    
Part_Handle_Err:
    MsgBox "There was an issue either setting the Part Number on the START HERE page, " & vbCrLf _
        & "Or we had an issue with splitting the Revision from the PartNum_Rev format" & vbCrLf & vbCrLf _
        & "Double check that the START HERE page is not protected"

End Sub

Public Sub partCombo_OnGetEnabled(ByRef control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = partCombo_Enabled
End Sub

Public Sub partCombo_OnGetItemCount(ByRef control As Office.IRibbonControl, ByRef Count As Variant)

    If Not Not partCombo_PartList Then
        Count = UBound(partCombo_PartList) + 1
    End If
End Sub

Public Sub partCombo_OnGetItemLabel(ByRef control As Office.IRibbonControl, ByRef index As Integer, ByRef ItemLabel As Variant)
    'After we initialize the partNumber list
    ItemLabel = partCombo_PartList(index)
End Sub

Public Sub partCombo_OnGetItemID(ByRef control As Office.IRibbonControl, ByRef index As Integer, ByRef ItemID As Variant)
    'Need to reference by ID? I guess not
End Sub

Public Sub partCombo_OnGetText(ByRef control As Office.IRibbonControl, ByRef Text As Variant)
    'when we don't have anything selected, use the placeholder text
    If partCombo_TextField <> vbNullString Then
        Text = partCombo_TextField
    Else
        Text = "[SELECT PART NUMBER]"
    End If

End Sub


'******************   RibbonMsg Label ***********************

Public Sub notificationLabel_OnGetLabel(ByRef control As Office.IRibbonControl, ByRef Label As Variant)
   Label = ribbonMsg
End Sub



'******************  Get API Key Button   ***********************
Public Sub GetAPIkey_OnAction(ByRef control As Office.IRibbonControl)
    HTTPconnections.AddCurrentUser
End Sub


'****************************************************
'**************   Dev Tools   ***********************
'****************************************************

'******************   Disable Events Btn  ***********************

Public Sub DisableEvents_Toggle(ByRef control As Office.IRibbonControl, ByRef isPressed As Boolean)
    Application.EnableEvents = Not (isPressed)
End Sub

'******************   Version History Btn  ***********************

Public Sub ShowVersionHistory(ByRef control As IRibbonControl)
'    MsgBox "show history"
    Load ChangeLogForm
    Dim changeLogText As String
    With ThisWorkbook.VBProject.VBComponents("DataSources").CodeModule
        changeLogText = (.Lines(23, .CountOfDeclarationLines))
    End With
    ChangeLogForm.changeLabel = changeLogText
    Debug.Print (Len(changeLogText))
'    ChangeLogForm.ScrollHeight = Len(changeLogText) / 8
    ChangeLogForm.ScrollHeight = ChangeLogForm.changeLabel.Height

    ChangeLogForm.Show
    Unload ChangeLogForm
End Sub

Public Sub GetVersionLabel(ByRef control As Office.IRibbonControl, ByRef Label As Variant)

   Label = "Version: " & DataSources.VERSION & vbCrLf & "Change History"
End Sub




'****************************************************
'***********   HelperFunctions   ********************
'****************************************************

Public Sub ClearComboVariables()
    partCombo_Enabled = False
    Erase partCombo_PartList
    partCombo_TextField = vbNullString
      
       
    InvalidateControl "notificationLabel"
    InvalidateControl "partCombo"

End Sub

Private Function GetPartIndex(partNum As String, json As Object) As Integer
    Dim i As Integer, part As Object
    i = 1
    For Each part In json
        Debug.Print (part("name"))
        If part("name") = partNum Then
            GetPartIndex = i
            Exit Function
        End If
        i = i + 1
    Next part

    Err.Raise Number:=vbObjectError + 5100, Description:="No Results returned by MeasurLink for this Part Number, Nothing to return"
End Function

Public Function GetParts_or_SetError() As String() ' ->  Returns Part Numbers with _Rev appended or Empty Array
    'Gets Part numbers from Variables and does some validation work required for the Custom Fields Group
    
    Dim rev As String
    rev = ThisWorkbook.Worksheets("START HERE").GetRevision()
    Dim partNums() As String
    partNums = ThisWorkbook.Worksheets("Variables").GetPartNumbers()
    
    'If there arent part numbers
    If (Not partNums) = -1 Then
        ResetViewControls notification_msg:="     Error:    No Part Numbers"
        Exit Function
        
    'If part numbers aren't unique
    ElseIf Not ThisWorkbook.Worksheets("Variables").IsUniquePartNumbers() Then
        ResetViewControls notification_msg:="     Error:    Non-Unique Part Numbers"
        Exit Function
    'If no Revision filled out on the START HERE page
    ElseIf rev = vbNullString Then
        ResetViewControls notification_msg:="     Error:    No Revision on START HERE"
        Exit Function
    'Otherwise its a valid list of partNumbers
    Else
        ClearRibbonNotification
        
        Dim i As Integer
        For i = 0 To UBound(partNums)
            partNums(i) = partNums(i) & "_" & rev  'Format it for PartNumbers in MeasurLink, must be PartNum_Rev
        Next i
        GetParts_or_SetError = partNums
    End If

End Function


Private Function Get_API_Key() As String '-> Returns contents of Key file or vbNullString
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    Dim user_path As String
    user_path = fso.BuildPath("C:\Users\", Environ("Username"))
    
    If user_path = vbNullString Then
        MsgBox "Couldn't Locate the Documents Directory for this User", vbCritical
        Exit Function
    End If
    
    user_path = fso.BuildPath(user_path, "Documents")
    user_path = fso.BuildPath(user_path, DataSources.API_KEY_FILE_NAME)
    
    If Not fso.FileExists(user_path) Then
        MsgBox "Couldn't locate an API_key file to supply," & vbCrLf & vbCrLf _
            & "Make sure to first request one with the 'Get API Key' Button" & vbCrLf _
            & "and follow instructions sent in the email. " & vbCrLf _
            & "See me if it is still happening", vbInformation
        Exit Function
    End If

    Get_API_Key = fso.OpenTextFile(user_path, ForReading).ReadAll()

End Function















