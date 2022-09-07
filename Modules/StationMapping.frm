VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StationMapping 
   Caption         =   "StationMapping"
   ClientHeight    =   11505
   ClientLeft      =   90
   ClientTop       =   255
   ClientWidth     =   8955.001
   OleObjectBlob   =   "StationMapping.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StationMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'*************************************************************
'*                  StationMapping
'*
'*  For all the Features in the given Routines that aren't mapped to stations yet,
'*      We will create these mappings, according to the QE's needs
'*************************************************************
'*************************************************************
Private listMaps() As Variant 'Representation of what is going on in our listBoxes
'listMaps  <--Parts Array
    'listMaps(0, i)  <-- PartName
    'listMaps(1, i)  <-- Part Selection Status
    'listMaps(2, i)  <-- Routines Array
        'listMaps(2, i)(0, j)  <-- Routine Name
        'listMaps(2, i)(1, j)  <-- Routine Selection Status
        'listMaps(2, i)(2, j)  <-- Features Array
            'listMaps(2, i)(2, j)(0, k)  <-- Feature Name
            'listMaps(2, i)(2, j)(1, k)  <-- Feature Type
            'listMaps(2, i)(2, j)(2, k)  <-- Mappings Array
                'listMaps(2, i)(2, j)(2, k)(0, m)  <-- Stations Name
                'listMaps(2, i)(2, j)(2, k)(1, m)  <-- DAQ Source Name
                'listMaps(2, i)(2, j)(2, k)(2, m)  <-- Status (-1, 0, 1) -> (exists already, don't add, add)

Private combo_resources() As Variant
Private combo_cells() As Variant
Private Const PLACEHOLDER = "Cell / Resource Combinations to Map"
Private active_rt As Variant
Private events_enabled As Boolean
Private cell_structure As Object
Public event_controls As Collection









'*************************************************************
'*************************************************************
'*                  Button CallBacks
'*************************************************************
'*************************************************************

'---------------   LoadOut Buttons  --------------------

Private Sub Add_Loadout_Button_Click()
    'Take the comboBox Values and Combine them into the LoadOutTextBox
    
    If Me.Loadout_TextBox.Text = PLACEHOLDER Then
        Me.Loadout_TextBox.Text = Me.CellComboBox.Value & "(" & Me.ResourceComboBox.Value & ")"
    Else
        Dim inp As String: inp = Me.CellComboBox.Value & "(" & Me.ResourceComboBox.Value & ")"
        Dim i As Integer: Dim mappings() As String
        mappings = Split(Me.Loadout_TextBox.Text, ",")
        For i = 0 To UBound(mappings)
            If mappings(i) = inp Then Exit Sub
        Next i
        Me.Loadout_TextBox.Text = Me.Loadout_TextBox.Text & "," & Me.CellComboBox.Value & "(" & Me.ResourceComboBox.Value & ")"
    End If

End Sub


Private Sub Apply_Loadout_Button_Click()
'Text the Values from our textBox and Map them in the Cells below
    
    If Me.Loadout_TextBox.Text = PLACEHOLDER Then Exit Sub
    
        
    'Validate we have Part and Routine(s) selected
    Dim part_ind As Integer, rt_inds() As Variant
    Dim i As Integer
    
    part_ind = Me.PartListBox.ListIndex
    If part_ind = -1 Then
        MsgBox "No Part Number Selected", vbCritical
    End If
    For i = 0 To UBound(Me.RoutineListBox.list)
        If Me.RoutineListBox.Selected(i) Then
            If (Not rt_inds) = -1 Then
                ReDim Preserve rt_inds(0)
                rt_inds(0) = i
            Else
                ReDim Preserve rt_inds(UBound(rt_inds) + 1)
                rt_inds(UBound(rt_inds)) = i
            End If
        End If
    Next i
    
    If (Not rt_inds) = -1 Then
        MsgBox "No Routines Selected", vbCritical
        Exit Sub
    End If
    
    Call ResetStationView
    
    'Extract our Station Mappings
    Dim re As RegExp: Set re = New RegExp: Dim match As Object
    re.Global = False
    re.Pattern = "(.*)\((.*)\)"
    
    
    Dim mappings() As String
    
    mappings = Split(Me.Loadout_TextBox.Text, ",")
    For i = 0 To UBound(mappings)
        Set match = re.Execute(mappings(i))
        On Error GoTo match_Error
        Dim cell As String: cell = match(0).SubMatches(0)
        Dim resource As String: resource = match(0).SubMatches(1)
        
        On Error GoTo 0
        
        For Each contr In Me.StationFrame.Controls(cell).Controls
            If TypeName(contr) = "CheckBox" And (contr.Tag = resource Or resource = "ALL") Then
                'contr.Value = True
                'Grab the station Name
                Dim station As String
                station = Split(contr.name, "check_")(1)
                
                MapToRoutines station:=station, rt_inds:=rt_inds, part_ind:=Me.PartListBox.ListIndex
                
            End If
        Next contr
    Next i
    
    
    Dim ft_list() As Variant
    ft_list = listMaps(2, Me.PartListBox.ListIndex)(2, active_rt)
    FillStationView feats_arr:=ft_list
    
    Me.Loadout_TextBox.Text = PLACEHOLDER
    Me.CellComboBox.Value = ""
    Me.ResourceComboBox.Value = ""
    
    Call EvalSelectionStatus
    
    Exit Sub
    
match_Error:
    MsgBox "We Couldn't parse the Loadout Here, Are you sure this Cell(Resource) mapping makes sense?", vbCritical

End Sub


Private Sub SubmitMappingsButton_Click()
    'TODO: move the code over from the settings button
    Dim json As Collection, fso As FileSystemObject, stream As TextStream
    Set json = Ravel()
    If json.Count = 0 Then Exit Sub
    
    Unload Me
    RibbonCommands.AddFeatureMapping json:=json
    
End Sub

Private Sub Loadout_Clear_Button_Click()
    Me.Loadout_TextBox = PLACEHOLDER
End Sub



'---------------   Part List Buttons  --------------------

Private Sub PartsButton_CopyMapping_Click()
    If Me.PartListBox.ListIndex = -1 Then
        MsgBox "There Must be a Part Number Selected!", vbInformation
        Exit Sub
    End If
    
    Dim rt_arr() As Variant
    rt_arr = listMaps(2, Me.PartListBox.ListIndex)
    If (Not rt_arr) = -1 Then
        MsgBox "The Selected Part must have one or more Routines", vbInformation
    End If
    
    If UBound(Me.PartListBox.list) = 0 Then
        MsgBox "There Must be More than a One Part Number", vbInformation
        Exit Sub
    End If
    
    'Build an array of Unique Station Names that we have as being mapped
    Dim i As Integer, j As Integer, k As Integer, m As Integer, stations() As Variant, routines() As Variant, features() As Variant, mappings() As Variant
    Dim changeRoutines() As Variant
    
    routines = listMaps(2, Me.PartListBox.ListIndex)
    ReDim Preserve changeRoutines(1, UBound(routines, 2))
    For j = 0 To UBound(routines, 2)
        changeRoutines(0, j) = routines(0, j)
    
        features = routines(2, j)
        If (Not features) = -1 Then GoTo cont_rt
        For k = 0 To UBound(features, 2)
            mappings = features(2, k)
            If (Not mappings) = -1 Then GoTo cont_features
            For m = 0 To UBound(mappings, 2)
                If mappings(2, m) = -1 Or mappings(2, m) = 0 Then GoTo cont_mappings
                If (Not stations) = -1 Then
                    ReDim Preserve stations(0)
                    stations(0) = mappings(0, m)
                ElseIf Not StationInArray(mappings(0, m), stations) Then
                    ReDim Preserve stations(UBound(stations) + 1)
                    stations(UBound(stations)) = mappings(0, m)
                End If
cont_mappings:
            Next m
cont_features:
        Next k
        
        changeRoutines(1, j) = stations
        Erase stations
cont_rt:
    Next j
    
    
    'For each other Part Number, iterate through the given Station Names,
    Dim rt_inds() As Variant
    For i = 0 To UBound(Me.PartListBox.list)
        If i = Me.PartListBox.ListIndex Then GoTo cont_parts
        routines = listMaps(2, i)
        If (Not routines) = -1 Then GoTo cont_parts
        
        'Iterate throught the names of the routines and find where we both
            'have a match of routine names and the stations array isnt nothing, otherwise continue
        For j = 0 To UBound(routines, 2)
            For k = 0 To UBound(changeRoutines, 2)
                If routines(0, j) = changeRoutines(0, k) Then
                    stations = changeRoutines(1, k)
                    If (Not stations) = -1 Then GoTo cont_rts
                    For m = 0 To UBound(stations)
                        Dim temp_rt() As Variant
                        temp_rt = Array(j)
                        MapToRoutines CStr(stations(m)), temp_rt, part_ind:=i
                    Next m
                End If
            Next k
            
cont_rts:
        Next j
cont_parts:
    Next i
    
    Call EvalSelectionStatus
    
End Sub

Private Function StationInArray(station_name As Variant, station_arr() As Variant) As Boolean
    Dim i As Integer
    For i = 0 To UBound(station_arr)
        If station_name = station_arr(i) Then
            StationInArray = True
            Exit Function
        End If
    Next i
End Function


'---------------   Routines List Buttons  --------------------

Private Sub RoutinesButton_ClearAll_Click()
    If IsNull(Me.RoutineListBox.list) Then Exit Sub
    
    On Error Resume Next
    Dim i As Integer
    For i = 0 To UBound(Me.RoutineListBox.list)
        Me.RoutineListBox.Selected(i) = False
    Next i
    On Error GoTo 0
    
    Toggle_CheckBoxes Enabled:=False
End Sub


Private Sub RoutinesButton_SelectAll_Click()
    If IsNull(Me.RoutineListBox.list) Then Exit Sub
    
    On Error Resume Next
    Dim i As Integer
    For i = 0 To UBound(Me.RoutineListBox.list)
        Me.RoutineListBox.Selected(i) = True
    Next i
    On Error GoTo 0
    
    Toggle_CheckBoxes Enabled:=True
End Sub




'---------------   Station Buttons  --------------------

Private Sub Stations_Settings_Button_Click()
    Load StationModify
    StationModify.Unravel_cells json:=cell_structure
    
    StationModify.Show vbModeless
    Me.Hide
End Sub

Private Sub Stations_Clear_Button_Click()
    If Me.PartListBox.ListIndex = -1 Then
        MsgBox "At Least a Part Number must be Selected", vbCritical
        Exit Sub
    End If
    
    Dim i As Integer, rt_inds() As Variant
    If Not IsNull(Me.RoutineListBox.list) Then
        For i = 0 To UBound(Me.RoutineListBox.list)
            If Me.RoutineListBox.Selected(i) Then
                If (Not rt_inds) = -1 Then
                    ReDim Preserve rt_inds(0)
                    rt_inds(0) = i
                Else
                    ReDim Preserve rt_inds(UBound(rt_inds) + 1)
                    rt_inds(UBound(rt_inds)) = i
                End If
            End If
        Next i
    End If
    
    'If we didnt have any routines selected, then we should clear everything from all of the routines for the Selected Part
    If (Not rt_inds) = -1 Then
        ReDim Preserve rt_inds(UBound(listMaps(2, Me.PartListBox.ListIndex), 2))
        For i = 0 To UBound(rt_inds)
            rt_inds(i) = i
        Next i
    End If
    
    Dim j As Integer, k As Integer, m As Integer
    For i = 0 To UBound(rt_inds)
        Dim ft_arr() As Variant, map_arr() As Variant
        ft_arr = listMaps(2, Me.PartListBox.ListIndex)(2, rt_inds(i))
        If (Not ft_arr) = -1 Then Exit Sub 'if there are not routines for the part
        For j = 0 To UBound(ft_arr, 2)
            'For each feature in routine, get its mappings
            map_arr = ft_arr(2, j)
            If (Not map_arr) = -1 Then GoTo cont_feat
            For k = 0 To UBound(map_arr, 2)
                'Iterate through the mappings, set all values to 0, unless they are -1
                If Not (map_arr(2, k) = -1) Then
                    map_arr(2, k) = 0
                End If
            Next k
            
            ft_arr(2, j) = map_arr
cont_feat:
        Next j
        
        '...Apply the feature array
        listMaps(2, Me.PartListBox.ListIndex)(2, rt_inds(i)) = ft_arr
        
        Erase ft_arr
    Next i
    
    EvalSelectionStatus
    ResetStationView

End Sub


'*************************************************************
'*************************************************************
'*                  ComboBox CallBacks
'*************************************************************
'*************************************************************


Private Sub CellComboBox_Change()
    

    Dim val As String: val = Me.CellComboBox.Value
    If val = vbNullString Then
        With Me.ResourceComboBox
            .Enabled = False
            .Value = ""
        End With
        Exit Sub
    
    End If
    
    'Enable the Resource Combo Box and Bring in its values
    Dim i As Integer
    For i = 0 To UBound(combo_cells)
        If val = combo_cells(i) Then
            With Me.ResourceComboBox
                .Enabled = True
                .Value = "ALL"
                .list = combo_resources(i)
            End With
            
        End If
    Next i
    
End Sub

Private Sub ResourceComboBox_Change()
    Me.Add_Loadout_Button.Enabled = Not Me.ResourceComboBox.Value = vbNullString
End Sub




Private Sub UserForm_Activate()
    Dim parts() As Variant, cols() As Variant
    cols = Array(0, 1)
    parts = SliceCols(listMaps, cols)
    parts = Application.Transpose(parts)
    parts = Force2D(parts)
    Me.PartListBox.list = parts
    Toggle_CheckBoxes Enabled:=False
    events_enabled = True
End Sub



'*************************************************************
'*************************************************************
'*                  ListBox CallBacks
'*************************************************************
'*************************************************************
Private Sub PartListBox_Change()
    Dim rts() As Variant, cols() As Variant
    cols = Array(0, 1)
    rts = listMaps(2, Me.PartListBox.ListIndex)
    rts = SliceCols(rts, cols)
    rts = Application.Transpose(rts)
    rts = Force2D(rts)
    
    Me.RoutineListBox.list = rts
    active_rt = missing
    Me.FeatureListBox.Clear
    Call ResetStationView
    
End Sub



Private Sub RoutineListBox_Change()
    Call ResetStationView
    
    If Not events_enabled Then Exit Sub

    Dim i As Integer, j As Integer, change_made As Boolean, ft_list() As Variant, cols() As Variant
    'Store the First selected Routine in Me.active_rt, erase it when it becomes unselected
    If IsEmpty(active_rt) Then
        For i = 0 To UBound(Me.RoutineListBox.list)
            If Me.RoutineListBox.Selected(i) Then
                active_rt = i
                change_made = True
            End If
        Next i
    ElseIf Not (Me.RoutineListBox.Selected(active_rt)) Then
        active_rt = missing
        Me.FeatureListBox.Clear
        For i = 0 To UBound(Me.RoutineListBox.list)
            If Me.RoutineListBox.Selected(i) Then
                active_rt = i
                change_made = True
            End If
        Next i
    End If
    
    If change_made Then
        'If a new routine has been set as the actgive_rt
        cols = Array(0, 1)
        ft_list = listMaps(2, Me.PartListBox.ListIndex)(2, active_rt)
        ft_list = SliceCols(ft_list, cols)
        ft_list = Application.Transpose(ft_list)
        ft_list = Force2D(ft_list) ' Tranpose will turn a single row into a 1-D array, need to turn it back
                
        For j = 1 To UBound(ft_list)
                'If its a variable feature type
            If ft_list(j, 2) = 1 Then
                ft_list(j, 2) = DataSources.ITEM_NOT_APPLICABLE & " / " & DataSources.ITEM_UNSELECTED
                'If its an attribute feature type
            Else
                ft_list(j, 2) = DataSources.ITEM_UNSELECTED & " / " & DataSources.ITEM_NOT_APPLICABLE
            End If
        Next j
        
        Me.FeatureListBox.list = ft_list
        
    End If
    
    If Not IsEmpty(active_rt) Then
        Toggle_CheckBoxes Enabled:=True
    
        Erase ft_list
        ft_list = listMaps(2, Me.PartListBox.ListIndex)(2, active_rt)
        FillStationView feats_arr:=ft_list
    End If
            
End Sub



Private Sub FeatureListBox_Change()
    ResetStationView skipCheckboxes:=True
    Dim maps() As Variant, upper_maps As Integer, i As Integer
    If Me.FeatureListBox.ListIndex = -1 Then Exit Sub
    
    'When we have a Part / Routine / Feature selected then
        'iterate through and visually set the stations that this feature is mapped to
    maps = listMaps(2, Me.PartListBox.ListIndex)(2, active_rt)(2, Me.FeatureListBox.ListIndex)
    If (Not maps) = -1 Then Exit Sub
    
    upper_maps = UBound(maps, 2)
    For Each contr In Me.StationFrame.Controls
            If TypeName(contr) = "Label" Then
                For i = 0 To upper_maps
                    If maps(0, i) = contr.Caption And maps(2, i) = -1 Then
                        contr.Font.Bold = True
                        contr.Parent.Font.Bold = True
                        GoTo cont_controls
                    End If
                Next i
            End If
cont_controls:
    Next contr

End Sub



'*************************************************************
'*************************************************************
'*                  Helper Functions
'*************************************************************
'*************************************************************

Private Sub ResetStationView(Optional skipCheckboxes As Boolean)
    events_enabled = False

    For Each contr In Me.StationFrame.Controls
            If TypeName(contr) = "Label" Or TypeName(contr) = "Frame" Then
                contr.Font.Bold = False
            ElseIf TypeName(contr) = "CheckBox" And Not skipCheckboxes Then
                contr.Value = False
            End If
    Next contr
    
    events_enabled = True
End Sub

'Helper Function
Private Sub FillStationView(feats_arr() As Variant)
    'Iterate through the Features of the selected Routine and set the Checkboxes or turn the labels bold
    
    events_enabled = False
    
    If (Not feats_arr) = -1 Then Exit Sub
    Dim maps() As Variant, i As Integer, j As Integer
    
    For i = 0 To UBound(feats_arr, 2)
        maps = feats_arr(2, i)
        If (Not maps) = -1 Then GoTo cont_feats
        For j = 0 To UBound(maps, 2)
            For Each contr In Me.StationFrame.Controls
                If TypeName(contr) = "CheckBox" Then
                    If contr.name = "check_" & maps(0, j) Then
                        'If maps(2, j) = -1 maps(2, j) = 1 Then
                        If maps(2, j) = 1 Then
                            contr.Value = True
                        End If
                    End If
                End If
            Next contr
            
        Next j
    
cont_feats:
        Erase maps
    Next i
    
    events_enabled = True
End Sub

'Called by Activate, RoutineListChange() and PartListChange()
Private Sub Toggle_CheckBoxes(Enabled As Boolean)
    For Each contr In Me.StationFrame.Controls
        If TypeName(contr) = "CheckBox" Then
            contr.Enabled = Enabled
        End If
    Next contr
End Sub



'Used by the ListBoxes
Private Function Force2D(arr() As Variant) As Variant()
    'If the array is 1-D becuase a single row was returned, convert it to 2D
    Dim i As Integer
    On Error Resume Next
    i = UBound(arr, 2)
    If Err.Number <> 0 Then
        Dim out(1 To 1, 1 To 2) As Variant
        out(1, 1) = arr(LBound(arr))
        out(1, 2) = arr(LBound(arr) + 1)
        Force2D = out
        Exit Function
    End If

    Force2D = arr

End Function

'Used to set the ListBoxes
Public Function SliceColumns(arr() As Variant, cols() As Variant)
    'Returns the (col2, n) array
    'This is needed because Applcation.Index() won't work with staggered Arrays best I can tell
    
    'Parameters
        'cols() -> Array(0,1,... n)
    On Error GoTo dimensional_err
    
    Dim col_size As Integer, row_size As Integer, i As Integer, j As Integer
    col_size = Application.Count(cols)
    row_size = UBound(arr, 2)
    
    Dim return_arr() As Variant
    ReDim Preserve return_arr(col_size - 1, row_size)
    
    For i = 1 To col_size
        For j = 0 To row_size
            return_arr(i - 1, j) = arr(cols(i - 1), j)
        Next j
    Next i
    
    SliceColumns = return_arr
    Exit Function
    
dimensional_err:
    MsgBox "The given Array is empty or the dimensions given don't match the dimensions of the Array", vbCritical
    SliceColumns = Array()

End Function

'Used by Unravel()
Private Function RoutineInArray(rt As Variant, part As Variant, arr() As Variant) As Integer
    'Return the index if found, or -1 if not found
    'feat_arr(0, i) -> RoutineName
    Dim i As Integer
    For i = 0 To UBound(arr, 2)
        If arr(0, i) = rt Then
            RoutineInArray = i
            Exit Function
        End If
    Next i
    
    RoutineInArray = -1
End Function

'Used by Unravel()
Private Function FeatureInArray(ft As Variant, arr() As Variant) As Integer
    'Return the index if found, or -1 if the index is not found
    'ftList(0, i) -> RoutineName
    Dim i As Integer
    For i = 0 To UBound(arr, 2)
        If arr(0, i) = ft Then
            FeatureInArray = i
            Exit Function
        End If
    Next i
    FeatureInArray = -1
End Function

'Used by Unravel()
Private Function IsArrayEmpty(arr()) As Boolean
    'Cant do the normal check for staggered arrays, need ot run into the error
    On Error Resume Next
    Dim i As Integer
    i = UBound(arr)
    
    IsArrayEmpty = Err.Number <> 0
    
    On Error GoTo 0
End Function



Private Function SliceCols(arr() As Variant, cols() As Variant)
    'Returns the (col2, n) array
    'This is needed because Applcation.Index() won't work with staggered Arrays best I can tell
    
    'Parameters
        'cols() -> Array(0,1,... n)
    On Error GoTo dimensional_err
    
    Dim col_size As Integer, row_size As Integer, i As Integer, j As Integer
    col_size = Application.Count(cols)
    row_size = UBound(arr, 2)
    
    Dim return_arr() As Variant
    ReDim Preserve return_arr(col_size - 1, row_size)
    
    For i = 1 To col_size
        For j = 0 To row_size
            return_arr(i - 1, j) = arr(cols(i - 1), j)
        Next j
    Next i
    
    SliceCols = return_arr
    Exit Function
    
dimensional_err:
    MsgBox "The given Array is empty or the dimensions given don't match the dimensions of the Array", vbCritical
    SliceCols = Array()

End Function




'*************************************************************
'*************************************************************
'*                  Public Callables
'*************************************************************
'*************************************************************


'Called by ThisWorkbook.Build_StationMappingForm()
Public Sub Set_Data(cells() As Variant, resources() As Variant, json_config As Object)
    combo_cells = cells
    combo_resources = resources
    Set cell_structure = json_config
End Sub

'Called by ThisWorkbook.Build_StationMappingForm()
Public Sub Unravel(json_content As Object)
    
    Dim i As Integer, j As Integer, k As Integer, upper_parts As Integer, upper_rts As Integer, upper_feats As Integer
    Dim feat_arr() As Variant, empT() As Variant
    
    For Each part In json_content
        If (Not listMaps) = -1 Then
            ReDim Preserve listMaps(2, 0)
            listMaps(0, 0) = part("name")
            listMaps(1, 0) = DataSources.ITEM_UNSELECTED
        Else
'            upper_parts = UBound(listMaps, 2) + 1
            upper_parts = upper_parts + 1
            ReDim Preserve listMaps(2, upper_parts)
            listMaps(0, upper_parts) = part("name")
            listMaps(1, upper_parts) = DataSources.ITEM_UNSELECTED
        End If
        
        Dim rtList() As Variant
        For Each feat In part("features")
            If (Not rtList) = -1 Then
                'First feat we add, just add all the routines the feature exists in, they will be unique anyways
                For i = 1 To feat("routines").Count
                    ReDim Preserve rtList(2, i - 1)
                    rtList(0, i - 1) = Split(feat("routines")(i)("name"), part("name") & "_")(1)  'asdf
                    rtList(1, i - 1) = DataSources.ITEM_UNSELECTED
                    
                    'TODO: we shuyold also be adding in teh features here
                        'As long as this doesnt conflict with what we are doing for the mappings below
                        
                    ReDim Preserve feat_arr(2, 0)
                    feat_arr(0, 0) = feat("name")
                    feat_arr(1, 0) = feat("type")
                    feat_arr(2, 0) = empT
                    rtList(2, i - 1) = feat_arr
                    
                    Erase feat_arr
                    
                Next i
            Else
                For i = 1 To feat("routines").Count
                    
                    Dim rt_ind As Integer
                    rt_ind = RoutineInArray(Split(feat("routines")(i)("name"), part("name") & "_")(1), part("name"), rtList)
                    If rt_ind = -1 Then
                        'If the routine doesnt already appear in our list
                        upper_rts = UBound(rtList, 2) + 1
                        ReDim Preserve rtList(2, upper_rts)
                        rtList(0, upper_rts) = Split(feat("routines")(i)("name"), part("name") & "_")(1)
                        rtList(1, upper_rts) = DataSources.ITEM_UNSELECTED
                        
                        'TODO: add in the feature here
                        ReDim Preserve feat_arr(2, 0)
                        feat_arr(0, 0) = feat("name")
                        feat_arr(1, 0) = feat("type")
                        feat_arr(2, 0) = empT
                        
                        rtList(2, upper_rts) = feat_arr
                        
                    Else
                        'we discovered the routine, so now we should add the feature to it
                        Dim upper_feat As Integer
                        feat_arr = rtList(2, rt_ind)   '<- get the index from above
                        upper_feat = UBound(feat_arr, 2) + 1
                        ReDim Preserve feat_arr(2, upper_feat)
                        feat_arr(0, upper_feat) = feat("name")
                        feat_arr(1, upper_feat) = feat("type")
                        feat_arr(2, upper_feat) = empT
                        
                        rtList(2, rt_ind) = feat_arr
                        
                    End If
                    
                    Erase feat_arr
                Next i
            End If
            
            'Start adding the features, make sure to check that the routineArray is not empty
            'We are only interested in adding features that we know exist in the Routines, regardless of whether they're mapped or not
            If (Not rtList) = -1 Then GoTo cont_feats
            
            
            For i = 0 To UBound(rtList, 2)
                For Each mapping In feat("mappings")
                    Dim maps() As Variant, feats() As Variant, rt_feats() As Variant
                
                    If rtList(0, i) <> Split(mapping("routine")("name"), part("name") & "_")(1) Then GoTo cont_maps
                    rt_feats = rtList(2, i)
                    If IsArrayEmpty(rt_feats) Then
                        ReDim Preserve feats(2, 0)
                        ReDim Preserve maps(2, 0)
                        feats(0, 0) = feat("name")
                        feats(1, 0) = feat("type")
                        
                        maps(0, 0) = mapping("daqsource")("station")("name")
                        maps(1, 0) = mapping("daqsource")("name")
                        maps(2, 0) = -1
                        
                        feats(2, 0) = maps
                        rtList(2, i) = feats
                    Else
                        'Need to find if the feature occurs in the existing array or not
                        Dim feat_ind As Integer
                        feat_ind = FeatureInArray(feat("name"), rt_feats)
                        
                        'If feature found
                        If Not feat_ind = -1 Then
                            'Should already have a daq source in it, so we need to append on
                            Dim upper_maps As Integer
                            maps = rtList(2, i)(2, feat_ind)
                            If (Not maps) = -1 Then
                                upper_maps = 0
                            Else
                                upper_maps = UBound(maps, 2) + 1
                            End If
                            ReDim Preserve maps(2, upper_maps)
                            maps(0, upper_maps) = mapping("daqsource")("station")("name")
                            maps(1, upper_maps) = mapping("daqsource")("name")
                            maps(2, upper_maps) = -1
                            rtList(2, i)(2, feat_ind) = maps
                            
                        Else 'else not found
                            'Add the feature and the first instance of a daq source to it
                            ReDim Preserve maps(2, 0)
                            maps(0, 0) = mapping("daqsource")("station")("name")
                            maps(1, 0) = mapping("daqsource")("name")
                            maps(2, 0) = -1
                            
                            upper_feats = UBound(rt_feats, 2) + 1
                            ReDim Preserve rt_feats(2, upper_feats)
                            rt_feats(0, upper_feats) = feat("name")
                            rt_feats(1, upper_feats) = feat("type")
                            rt_feats(2, upper_feats) = maps
                            
                            rtList(2, i) = rt_feats
                        End If
                    End If
                    
                    Erase maps
                    Erase feats
                    Erase feat_arr
cont_maps:
                Next mapping
            Next i
cont_feats:
        Next feat
        
        'Add the Routines to the Top-Level part
        listMaps(2, upper_parts) = rtList
        Erase rtList
    Next part
    
    
End Sub

'Called by ThisWorkbook.Build_StationMappingForm()
Public Function Ravel() As Collection
    'Turn the listMaps() array back into a JSON format.
    'Remember to skip anything that isnt marked as being changed, and anything that has 0 or -1 at the value for the station
    Set parts_out = New Collection
    
    Dim i As Integer, j As Integer, k As Integer, m As Integer
    
    For i = 0 To UBound(Me.PartListBox.list)
        If Not (listMaps(1, i) = DataSources.ITEM_SELECTED) Then GoTo cont_parts_selected
        Set temp_part = New Dictionary: Set temp_features = New Collection
        temp_part.Add "name", listMaps(0, i)
        temp_part.Add "features", temp_features
        
        Dim routines() As Variant, features() As Variant, mappings() As Variant
        routines = listMaps(2, i)
        If (Not routines) = -1 Then GoTo cont_parts_selected
        For j = 0 To UBound(routines, 2)
            If Not (routines(1, j) = DataSources.ITEM_SELECTED) Then GoTo cont_rts_selected
            
            features = routines(2, j)
            If (Not features) = -1 Then GoTo cont_rts_selected
            For k = 0 To UBound(features, 2)
                mappings = features(2, k)
                If (Not mappings) = -1 Then GoTo cont_fts
                
                For m = 0 To UBound(mappings, 2)
                    If mappings(2, m) = 1 Then
                        Dim feats_copy As Collection
                        Set feats_copy = temp_part("features")
                        
                        Set mapping_dict = New Dictionary: Set map_rt = New Dictionary: Set map_daq = New Dictionary: Set map_station = New Dictionary
                        Dim temp_feat As Dictionary, mapping_coll As Collection, feat_ind As Integer
                        feat_ind = GetFeatIndFromColl(features(0, k), feats_copy)
                        If feat_ind = 0 Then
                            'Haven't added this feature yet, so Lets add it
                            Set temp_feat = New Dictionary: Set mapping_coll = New Collection
                            temp_feat.Add "name", features(0, k)
                            temp_feat.Add "mappings", mapping_coll
                        Else
                            'Otherwise the Feature and Mappings exist already
                            Set temp_feat = feats_copy(feat_ind)
                            Set mapping_coll = temp_feat("mappings")
                        End If
                        
                        Set map_rt = New Dictionary
                        map_rt.Add "name", listMaps(0, i) & "_" & routines(0, j)
                        
                        Set map_daq = New Dictionary
                        Set map_station = New Dictionary
                        
                        map_station.Add "name", mappings(0, m)
                        If features(1, k) = 1 Then 'If its a Variable Feature
                            map_daq.Add "name", "Keyboard"
                        Else  'If its an Attribute Feature
                            map_daq.Add "name", "Pass/Fail"
                        End If
                        
                        map_daq.Add "station", map_station
                        
                        mapping_dict.Add "routine", map_rt
                        mapping_dict.Add "daqsource", map_daq
                        
                        mapping_coll.Add mapping_dict
                        
                        Set temp_feat.item("mappings") = mapping_coll
                        
                        'If this is the first time we created this feature, no need to remove anything
                        If feat_ind <> 0 Then
                            feats_copy.Remove index:=feat_ind
                        End If
                        
                        feats_copy.Add temp_feat
                        
                        'Add back in our revised Collection of Features
                        Set temp_part.item("features") = feats_copy
                    
                    End If
                Next m
                
cont_fts:
            Next k
            
        
cont_rts_selected:
        Next j

        'If the part dictionary has features in it, then add it into parts_out
        If temp_part("features").Count <> 0 Then
            parts_out.Add temp_part
        End If
cont_parts_selected:
    Next i
    
    
    Set Ravel = parts_out   'have recipient of Ravel() check if the Collection is empty of Not
    
End Function

'TODO: Move to helper() functions
'Called By Ravel()
Private Function GetFeatIndFromColl(feat_name As Variant, coll As Collection) As Integer
    If coll.Count = 0 Then Exit Function
    Dim i As Integer
    For i = 1 To coll.Count
        If coll(i)("name") = feat_name Then
            GetFeatIndFromColl = i
            Exit Function
        End If
    Next i
    
    GetFeatFromColl = 0
End Function



'Called by Apply_Loadout_Button_Click and Change events of checkBoxClass instances
Public Sub MapToRoutines(station As String, rt_inds() As Variant, Optional part_ind As Variant, Optional checkBox_Hook As Boolean, Optional checkBox_Value As Boolean) 'rt_inds must be Array of Integer Indexes

    'If a checkbox is being changed by a sub in this Userform, then ignore the callback
    If checkBox_Hook And Not events_enabled Then Exit Sub
    
    'If this was called by a checkbox, then we need to set this here
    If IsMissing(part_ind) Then part_ind = Me.PartListBox.ListIndex
        
    'If its a callback from a checkBoxClass, then we won't already have the list of routines collected
    If (Not rt_inds) = -1 Then
        Dim i As Integer
        For i = 0 To UBound(Me.RoutineListBox.list)
            If Me.RoutineListBox.Selected(i) Then
                If (Not rt_inds) = -1 Then
                    ReDim Preserve rt_inds(0)
                    rt_inds(0) = i
                Else
                    ReDim Preserve rt_inds(UBound(rt_inds) + 1)
                    rt_inds(UBound(rt_inds)) = i
                End If
            End If
        Next i
    End If
    'TODO: need to build out the array of the seleced routines

    Dim j As Integer, k As Integer, m As Integer, feats_arr() As Variant, maps_arr() As Variant
    
'For the selected part, for each routine, for each feature in the routine...
    For j = 0 To UBound(rt_inds)
        feats_arr = listMaps(2, part_ind)(2, rt_inds(j))
        For k = 0 To UBound(feats_arr, 2)
            maps_arr = feats_arr(2, k)
            'If we dont have any mappings, then we should add one here,
            If (Not maps_arr) = -1 Then
                ReDim Preserve maps_arr(2, 0)
                maps_arr(0, 0) = station
                If feats_arr(1, k) = 1 Then
                    maps_arr(1, 0) = "Keyboard"
                Else
                    maps_arr(1, 0) = "Pass/Fail"
                End If
                If (Not checkBox_Hook) Or (checkBox_Value = True) Then
                    maps_arr(2, 0) = 1
                Else
                    maps_arr(2, 0) = 0
                End If
            Else 'Otherwise, we need iterate till we either find the station, or realize we need to append on a station
                For m = 0 To UBound(maps_arr, 2)
                    If maps_arr(0, m) = station Then
                        If maps_arr(2, m) <> -1 Then
                            If (Not checkBox_Hook) Or (checkBox_Value = True) Then
                                maps_arr(2, m) = 1
                            Else
                                maps_arr(2, m) = 0
                            End If
                        End If
                        GoTo cont_feats
                    End If
                Next m
                
                'we didnt find the station so we need to add it
                Dim upp_maps As Integer
                upp_maps = UBound(maps_arr, 2) + 1
                ReDim Preserve maps_arr(2, upp_maps)
                maps_arr(0, upp_maps) = station
                If feats_arr(1, k) = 1 Then
                    maps_arr(1, upp_maps) = "Keyboard"
                Else
                    maps_arr(1, upp_maps) = "Pass/Fail"
                End If
                If (Not checkBox_Hook) Or (checkBox_Value = True) Then
                    maps_arr(2, upp_maps) = 1
                Else
                    maps_arr(2, upp_maps) = 0
                End If
            End If
            

cont_feats:
            'Append the new mappings back on the feature
            feats_arr(2, k) = maps_arr
            
            Erase maps_arr
        Next k
        'Append the new feature back on the routine
        listMaps(2, part_ind)(2, rt_inds(j)) = feats_arr
        
        Erase feats_arr
    Next j

End Sub

'Called by ApplyLoadoutButton and change events of checkBoxClass intances
Public Sub EvalSelectionStatus(Optional checkBox_Hook As Boolean)
    If checkBox_Hook And Not events_enabled Then Exit Sub

    Dim i As Integer, j As Integer, k As Integer, m As Integer
    Dim rt_arr() As Variant, ft_arr() As Variant, maps_arr() As Variant
    
    For i = 0 To UBound(listMaps, 2)
        rt_arr = listMaps(2, i)
        If (Not rt_arr) = -1 Then GoTo cont_parts
        
        For j = 0 To UBound(rt_arr, 2)
            ft_arr = rt_arr(2, j)
            For k = 0 To UBound(ft_arr, 2)
                maps_arr = ft_arr(2, k)
                If (Not maps_arr) = -1 Then GoTo cont_feats
                For m = 0 To UBound(maps_arr, 2)
                    If maps_arr(2, m) = 1 Then
                        listMaps(2, i)(1, j) = DataSources.ITEM_SELECTED
                        listMaps(1, i) = DataSources.ITEM_SELECTED
                        If i = Me.PartListBox.ListIndex Then Me.RoutineListBox.list(j, 1) = DataSources.ITEM_SELECTED
                        Me.PartListBox.list(i, 1) = DataSources.ITEM_SELECTED
                        GoTo cont_rts
                    End If
                Next m
cont_feats:
            Next k
            'Presumably we got here because nothing was mapped, so set the selection status as not_selected
            listMaps(2, i)(1, j) = DataSources.ITEM_UNSELECTED
            If Me.PartListBox.ListIndex = i Then Me.RoutineListBox.list(j, 1) = DataSources.ITEM_UNSELECTED
cont_rts:
        Next j
        'Check, if no routines are selected then we need to unselect the part
        For j = 0 To UBound(rt_arr, 2)
            If listMaps(2, i)(1, j) = DataSources.ITEM_SELECTED Then GoTo cont_parts 'Found one selected, so skip the rest
        Next j
        listMaps(1, i) = DataSources.ITEM_UNSELECTED
        Me.PartListBox.list(i, 1) = DataSources.ITEM_UNSELECTED
cont_parts:
    Next i
    
'ReDraw the Listboxes to show Selection Status
    'Normally I would just slice the col
    
    
End Sub











