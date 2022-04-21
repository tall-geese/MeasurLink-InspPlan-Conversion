VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} INSERTform 
   Caption         =   "Insert Custom Field Information"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10230
   OleObjectBlob   =   "INSERTform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "INSERTform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*************************************************************
'*************************************************************
'*                  INSERT Form
'*
'*************************************************************
'*************************************************************

Private json_routine_map As Object
Private listen_clk_events As Boolean
Private listParts() As Variant 'Representation of what is going on in our listBoxes
'listParts  <--Parts Array
    'listParts(0, i)  <-- PartName
    'listParts(1, i)  <-- list_selection_status
    'listParts(2, i)  <-- Feautres Array
        'listParts(2, i)(0, j)  <-- FeatureName
        'listParts(2, i)(1, j)  <-- list_selection_status
        'listParts(2, i)(2, j)  <-- CustomFields Array
            'listFields(2, i)(2, j)(0, k)  <-- Custom Field ID
            'listFields(2, i)(2, j)(1, k)  <-- list_selection_status


'************************************************************
'*********   Initialize / Exposed Routines   ****************
'************************************************************



    'Called by RibbonCommands.AddCustomFields_OnAction   after Load'ing this form
Public Sub BuildListArray(json_parts_api As Object, json_parts_map As Object)
    On Error GoTo buildListArrErr

    Set json_routine_map = json_parts_map
    
    If json_parts_map.Count = 0 Or json_parts_api.Count = 0 Then
        Err.Raise Number:=vbObjectError + 7000, Description:="None of the parts in the Variables tab exist in the Database yet!" & vbCrLf _
                        & "You must create QIF files first and Import them before you can Upload Custom Field information"
    End If
    
    'Go through the Part Names, everything we want to insert should have a matching part returned
    Dim part_dict_rm As Object, part_dict_api As Object
    For Each part_dict_rm In json_parts_map
        For Each part_dict_api In json_parts_api
            If part_dict_api("name") = part_dict_rm("name") Then
                If (Not listParts) = -1 Then
                    ReDim Preserve listParts(0 To 2, 0)
                Else
                    ReDim Preserve listParts(0 To 2, UBound(listParts, 2) + 1)
                End If
                
                'Set the values
                listParts(0, UBound(listParts, 2)) = part_dict_rm("name")
                listParts(1, UBound(listParts, 2)) = DataSources.ITEM_SELECTED
                GoTo next_map_part
            End If
        Next part_dict_api
next_map_part:
    Next part_dict_rm
    
    'If we made it here without adding any parts...
    If (Not listParts) = -1 Then
        Err.Raise Number:=vbObjectError + 7000, Description:="Couldn't find any of the Variables tab parts" & vbCrLf _
                    & "in the list of parts returned by the server."
    End If
    
    'TODO: need to check for null type instances of features, although this is highly unlikely if we made it this far...
    
    'Build Feature Maps
    Dim feat_dict_rm As Object, feat_dict_api As Object, i As Integer
    For i = 0 To UBound(listParts, 2)
        Dim listFeatures() As Variant
    
        For Each feat_dict_rm In json_parts_map(GetPartIndex(json_parts_map, listParts(0, i)))("features")
            For Each feat_dict_api In json_parts_api(GetPartIndex(json_parts_api, listParts(0, i)))("features")
                If feat_dict_api("name") = feat_dict_rm("name") Then
                    If (Not listFeatures) = -1 Then
                        ReDim Preserve listFeatures(0 To 2, 0)
                    Else
                        ReDim Preserve listFeatures(0 To 2, UBound(listFeatures, 2) + 1)
                    End If
                    listFeatures(0, UBound(listFeatures, 2)) = feat_dict_rm("name")
                    listFeatures(1, UBound(listFeatures, 2)) = DataSources.ITEM_SELECTED
                                        
                    GoTo next_map_feat
                End If
            Next feat_dict_api
                'got here without finding a match
                If (Not listFeatures) = -1 Then
                    ReDim Preserve listFeatures(0 To 2, 0)
                Else
                    ReDim Preserve listFeatures(0 To 2, UBound(listFeatures, 2) + 1)
                End If
                listFeatures(0, UBound(listFeatures, 2)) = feat_dict_rm("name")
                listFeatures(1, UBound(listFeatures, 2)) = DataSources.ITEM_NOT_FOUND
next_map_feat:
        Next feat_dict_rm
        
        listParts(2, i) = listFeatures
        Erase listFeatures
    Next i
    
    'TODO: check for existence of null values
    
    'Build Custom Field Maps
    Dim j As Integer
    
    For i = 0 To UBound(listParts, 2) 'For each part
        Set feat_dict_rm = json_parts_map(GetPartIndex(json_parts_map, listParts(0, i)))("features")
        Set feat_dict_api = json_parts_api(GetPartIndex(json_parts_api, listParts(0, i)))("features")
    
        For j = 0 To UBound(listParts(2, i), 2) 'For each feature
            Dim listFields(0 To 1, 0 To 6) As Variant
            listFields(0, 0) = 13
            listFields(0, 1) = 16
            listFields(0, 2) = 8
            listFields(0, 3) = 12
            listFields(0, 4) = 11
            listFields(0, 5) = 15
            listFields(0, 6) = 3
            
            Dim k As Integer
            For k = 0 To UBound(listFields, 2)
                listFields(1, k) = DataSources.ITEM_NOT_APPLICABLE
            Next k
            
                'If the feature wasnt in the db, then we can't add anything
            If listParts(2, i)(1, j) = DataSources.ITEM_NOT_FOUND Then
               listParts(2, i)(2, j) = listFields
               GoTo nextfeat
            End If
            
            'we need to get the dict of the features on both sides, check for the count.
            Dim feat_rm As Object, feat_api As Object
            Set feat_rm = feat_dict_rm(GetFeatureIndex(feat_dict_rm, listParts(2, i)(0, j)))
            Set feat_api = feat_dict_api(GetFeatureIndex(feat_dict_api, listParts(2, i)(0, j)))
            
            'TODO: test this block
            If feat_rm("custom_fields").Count = 0 Then  'We don't have anything we want to add
                listParts(2, i)(1, j) = DataSources.ITEM_NOT_FOUND
                
                For k = 0 To UBound(listFields, 2)
                    listFields(1, k) = DataSources.ITEM_NOT_FOUND
                Next k
                
                listParts(2, i)(2, j) = listFields
                GoTo nextfeat
            End If
            
            'This is the bread and butter, db has no custom fields and we want to add everything  we can
            If feat_api("custom_fields").Count = 0 Then
                Dim cf As Object
                For Each cf In feat_rm("custom_fields")
                    For k = 0 To UBound(listFields, 2)
                        If cf("customFieldId") = listFields(0, k) Then
                            listFields(1, k) = DataSources.ITEM_SELECTED
                        End If
                    Next k
                Next cf
                
                listParts(2, i)(2, j) = listFields
                
                GoTo nextfeat
            End If
            
            'All the custom fields for this feature exist already...
            If feat_api("custom_fields").Count = 7 Then
                listParts(2, i)(1, j) = DataSources.ITEM_NOT_APPLICABLE
                listParts(2, i)(2, j) = listFields
                GoTo nextfeat
            End If
    
            'Otherwise, there are some customFields we can upload and some we can't becuase they are occupied
            Dim cf_rm As Object, cf_api As Object
            For Each cf_rm In feat_rm("custom_fields")
                For Each cf_api In feat_api("custom_fields")
                    If cf_rm("customFieldId") = cf_api("customFieldId") Then
                        'Found a match, that means that we cant add anything. Leave status as [---]
                        GoTo next_cf
                    End If
                Next cf_api
                
                'No match found, change status so it can be Inserted
                For k = 0 To UBound(listFields, 2)
                    If cf_rm("customFieldId") = listFields(0, k) Then
                        listFields(1, k) = DataSources.ITEM_SELECTED
                    End If
                Next k
next_cf:
            
            Next cf_rm
            
            listParts(2, i)(2, j) = listFields
    
nextfeat:
        Next j
        
        'If all of the Feature for this part of an NA_type, then we need to make the Part NA as well
        Dim feat_count As Integer, na_feats As Integer
        feat_count = UBound(listParts(2, i), 2) + 1
        
        For j = 0 To UBound(listParts(2, i), 2)
            If listParts(2, i)(1, j) = DataSources.ITEM_NOT_APPLICABLE Or listParts(2, i)(1, j) = DataSources.ITEM_NOT_FOUND Then
                na_feats = na_feats + 1
            End If
        Next j
        
        If feat_count = na_feats Then
            listParts(1, i) = DataSources.ITEM_NOT_APPLICABLE
        End If
        
        na_feats = 0
        
    Next i

    'End with a call to Init_Lists
    init_lists
'    ResetLists init_lists:=True
    
    Exit Sub
    
buildListArrErr:
    MsgBox "Encountered a problem at INSERTform.BuildListArray()" & vbCrLf & Err.Description, vbCritical
    
    
    
End Sub

    Private Function GetPartIndex(parts As Object, partName As Variant) As Integer
        Dim i As Integer, part As Object
        i = 1
        For Each part In parts
            If part("name") = partName Then
                GetPartIndex = i
                Exit Function
            End If
            i = i + 1
        Next part
    End Function
    
    Private Function GetFeatureIndex(features As Object, featName As Variant) As Integer
        Dim i As Integer, feat As Object
        i = 1
        For Each feat In features
                If feat("name") = featName Then
                    GetFeatureIndex = i
                    Exit Function
                End If
            i = i + 1
        Next feat
    End Function


Private Sub init_lists()
    
    ReLoadParts
    ReLoadFeatures partIndex:=0
    ReLoadFields partIndex:=0, featureIndex:=0
    
    Me.PartListBox.Selected(0) = True
    Me.FeatureListBox.Selected(0) = True
    
    
    listen_clk_events = True
    
    Exit Sub
    
initListErr:
    MsgBox "Encountered Error at INSERTform.Init_Lists()" & vbCrLf & Err.Description, vbCritical
    
End Sub

    '************************************************************
    '****************   Helper Functions   **********************
    '************************************************************
    Private Sub ResetLists(Optional init_lists As Boolean)
        'Should be called whenever a change to the underlying array is made,
            'Not when something like a selection change occurs
        Dim part_ind As Integer, feat_ind As Integer
        
        If Not init_lists Then  'For the INIT, just leave the indexes as 0
            part_ind = Me.PartListBox.ListIndex
            feat_ind = Me.FeatureListBox.ListIndex
        End If
        
        ReLoadParts
        ReLoadFeatures partIndex:=part_ind
        ReLoadFields partIndex:=part_ind, featureIndex:=feat_ind
        
        listen_clk_events = False
        Me.PartListBox.Selected(part_ind) = True
        Me.FeatureListBox.Selected(feat_ind) = True
        listen_clk_events = True
           
    
    End Sub


    Private Sub ReLoadParts()
        Me.PartListBox.Clear
    
        Dim i As Integer
        For i = 0 To UBound(listParts, 2)
            Me.PartListBox.AddItem
            Me.PartListBox.List(i, 0) = listParts(0, i)
            Me.PartListBox.List(i, 1) = listParts(1, i)
        Next i
        
    End Sub

    Private Sub ReLoadFeatures(partIndex As Integer)
        Me.FeatureListBox.Clear
    
        Dim j As Integer
        For j = 0 To UBound(listParts(2, partIndex), 2)
            Me.FeatureListBox.AddItem
            Me.FeatureListBox.List(j, 0) = listParts(2, partIndex)(0, j)
            Me.FeatureListBox.List(j, 1) = listParts(2, partIndex)(1, j)
        Next j
    End Sub
    
    Private Sub ReLoadFields(partIndex As Integer, featureIndex As Integer)
        Me.FieldsListBox.Clear
    
        Dim k As Integer
        Me.FieldsListBox.AddItem
        For k = 0 To UBound(listParts(2, partIndex)(2, featureIndex), 2)
            Me.FieldsListBox.List(0, k) = listParts(2, partIndex)(2, featureIndex)(1, k)
        Next k
    End Sub
    
    
    Private Sub CF_Change(Optional user_select As Boolean)
           
        'Propogate up to Feature to Evaluate
        Feat_Change user_select:=user_select
    End Sub
    
    
    Private Sub Feat_Change(Optional user_select As Boolean)
        Dim feat_select_status As String, k As Integer
        feat_select_status = listParts(2, Me.PartListBox.ListIndex)(1, Me.FeatureListBox.ListIndex)
        
        'If the Feature is Unselected, then all applicable CustomFields should be unselected as well
        If feat_select_status = DataSources.ITEM_UNSELECTED Then
            For k = 0 To UBound(listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex), 2)
                If listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, k) = DataSources.ITEM_SELECTED Then
                    listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, k) = DataSources.ITEM_UNSELECTED
                End If
            Next k
            GoTo propChange 'If its the last feature to be unselected, then the part must be unselected as well
        
        'Otherwise, if the Feature is selected but None of the Custom Fields are, then is should become Unselected
        ElseIf feat_select_status = DataSources.ITEM_SELECTED And Not user_select Then
            Dim cf_select_count As Integer
            For k = 0 To UBound(listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex), 2)
                If listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, k) = DataSources.ITEM_SELECTED Then
                    cf_select_count = cf_select_count + 1
                End If
            Next k
        
            'If no custom_fields are selected, the feature should be unselected
            If cf_select_count = 0 Then listParts(2, Me.PartListBox.ListIndex)(1, Me.FeatureListBox.ListIndex) = DataSources.ITEM_UNSELECTED
        
        End If
        
propChange:
        'Propogate up to Part to Evaluate
        
        If Not user_click Then
            Part_Change
        End If
    
    End Sub
    
    Private Sub Part_Change(Optional user_select As Boolean)
        Dim part_select_status As String, j As Integer, k As Integer
        part_select_status = listParts(1, Me.PartListBox.ListIndex)
        
        'If a part is unselected, then all features and custom fields should be Unselected
        If part_select_status = DataSources.ITEM_UNSELECTED Then
            For j = 0 To UBound(listParts(2, Me.PartListBox.ListIndex), 2)
                If listParts(2, Me.PartListBox.ListIndex)(1, j) = DataSources.ITEM_SELECTED Then
                    listParts(2, Me.PartListBox.ListIndex)(1, j) = DataSources.ITEM_UNSELECTED
                End If
                
                For k = 0 To UBound(listParts(2, Me.PartListBox.ListIndex)(2, j), 2)
                    If listParts(2, Me.PartListBox.ListIndex)(2, j)(1, k) = DataSources.ITEM_SELECTED Then
                        listParts(2, Me.PartListBox.ListIndex)(2, j)(1, k) = DataSources.ITEM_UNSELECTED
                    End If
                Next k
            Next j
            
        'Otherwise, if the part is Selected, but we dont have any features selected in it,
        ElseIf part_select_status = DataSources.ITEM_SELECTED And Not user_select Then
            Dim feats_selected As Integer
            For j = 0 To UBound(listParts(2, Me.PartListBox.ListIndex), 2)
                If listParts(2, Me.PartListBox.ListIndex)(1, j) = DataSources.ITEM_SELECTED Then
                    feats_selected = feats_selected + 1
                End If
            Next j
            
            If feats_selected = 0 Then listParts(1, Me.PartListBox.ListIndex) = DataSources.ITEM_UNSELECTED
        
        End If
    
        'Finally make a call to the sub that reloads the listBoxes
        ResetLists
    
    End Sub










'************************************************************
'**************   ListBoxes Callbacks   *********************
'************************************************************


    '************************************************************
    '******************   PartListBox   *************************
    '************************************************************

Private Sub PartListBox_Click()
    If Not listen_clk_events Then Exit Sub
        
    ReLoadFeatures partIndex:=Me.PartListBox.ListIndex
    ReLoadFields partIndex:=Me.PartListBox.ListIndex, featureIndex:=0
    
    listen_clk_events = False
    Me.FeatureListBox.Selected(0) = True
    listen_clk_events = True
    
End Sub

Private Sub PartListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not listen_clk_events Then Exit Sub
    
    Dim part_val As String
    part_val = listParts(1, Me.PartListBox.ListIndex)
    
    'TODO: get rid of these
    If Not (part_val = DataSources.ITEM_SELECTED Or part_val = DataSources.ITEM_UNSELECTED) Then Exit Sub
    
    Dim index As Integer
    index = Me.PartListBox.ListIndex
    
    If part_val = DataSources.ITEM_SELECTED Then
        listParts(1, index) = DataSources.ITEM_UNSELECTED
    Else
        listParts(1, index) = DataSources.ITEM_SELECTED
    End If
    
    Part_Change user_select:=True
    
End Sub

        '************************************************************
        '****************   FeatureListBox   ************************
        '************************************************************

Private Sub FeatureListBox_Click()
    If Not listen_clk_events Then Exit Sub
    
    ReLoadFields partIndex:=Me.PartListBox.ListIndex, featureIndex:=Me.FeatureListBox.ListIndex
        
End Sub

Private Sub FeatureListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Not listen_clk_events Then Exit Sub
    
    Dim feat_val As String
    feat_val = listParts(2, Me.PartListBox.ListIndex)(1, Me.FeatureListBox.ListIndex)
    
    If Not (feat_val = DataSources.ITEM_SELECTED Or feat_val = DataSources.ITEM_UNSELECTED) Then Exit Sub
    
    Dim index As Integer
    index = Me.FeatureListBox.ListIndex
    
    If feat_val = DataSources.ITEM_SELECTED Then
        listParts(2, Me.PartListBox.ListIndex)(1, index) = DataSources.ITEM_UNSELECTED
    Else
        listParts(2, Me.PartListBox.ListIndex)(1, index) = DataSources.ITEM_SELECTED
    End If
    
    Feat_Change user_select:=True

End Sub








        '************************************************************
        '***************   CustomField ListBox   ********************
        '************************************************************


Private Sub Balloon_Button_Click()
    Dim cf_val As String, cf_index As Integer
    cf_index = 0
    cf_val = listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, cf_index)
    
    
    switch_cf_and_reload select_status:=cf_val, cf_index:=cf_index
End Sub

Private Sub Description_Button_Click()
    Dim cf_val As String, cf_index As Integer
    cf_index = 1
    cf_val = listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, cf_index)
    
    
    switch_cf_and_reload select_status:=cf_val, cf_index:=cf_index
End Sub

Private Sub Comments_Button_Click()
    Dim cf_val As String, cf_index As Integer
    cf_index = 2
    cf_val = listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, cf_index)
    
    
    switch_cf_and_reload select_status:=cf_val, cf_index:=cf_index
End Sub

Private Sub Frequency_Button_Click()
    Dim cf_val As String, cf_index As Integer
    cf_index = 3
    cf_val = listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, cf_index)
    
    
    switch_cf_and_reload select_status:=cf_val, cf_index:=cf_index
End Sub

Private Sub Insp_Method_Button_Click()
    Dim cf_val As String, cf_index As Integer
    cf_index = 4
    cf_val = listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, cf_index)
    
    
    switch_cf_and_reload select_status:=cf_val, cf_index:=cf_index
End Sub



Private Sub Pins_Gauges_Button_Click()
    Dim cf_val As String, cf_index As Integer
    cf_index = 5
    cf_val = listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, cf_index)
    
    
    switch_cf_and_reload select_status:=cf_val, cf_index:=cf_index
End Sub

Private Sub Attr_Tol_Button_Click()
    Dim cf_val As String, cf_index As Integer
    cf_index = 6
    cf_val = listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, cf_index)
    
    
    switch_cf_and_reload select_status:=cf_val, cf_index:=cf_index
End Sub



    Private Sub switch_cf_and_reload(select_status As String, cf_index As Integer)
        If select_status = DataSources.ITEM_SELECTED Then
            listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, cf_index) = DataSources.ITEM_UNSELECTED
        ElseIf select_status = DataSources.ITEM_UNSELECTED Then
            listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, cf_index) = DataSources.ITEM_SELECTED
        Else
            Exit Sub
        
        End If
        
        CF_Change user_select:=True
        
    
    End Sub









'************************************************************
'*******************   GUI Buttons   ************************
'************************************************************



Private Sub HelpButton_Click()
    MsgBox "The Values in this form can be 1 of 4 possible states" & vbCrLf & vbCrLf _
        & "The Following 2 Values can be changed by double clicking the list item." & vbCrLf & vbCrLf _
        & "[ * ] - The Value is Selected to be added to MeasurLink." & vbCrLf _
        & "[   ] - The Value Not Selected and will not be added to MeasurLink" & vbCrLf _
        & vbCrLf _
        & "The Following 2 Values cannot be changed" & vbCrLf & vbCrLf _
        & "[---] - The MeasurLink database already has a value for this custom field, or all fields for this feature. You can't insert one." & vbCrLf _
        & "[???] - A Value can be inserted into the databse, but we couldn't find one on the PartLib Table.", vbInformation

End Sub



Private Sub PartListRefresh_Click()
    Dim i As Integer
    For i = 0 To UBound(listParts, 2)
        If listParts(1, i) = DataSources.ITEM_UNSELECTED Then
            listParts(1, i) = DataSources.ITEM_SELECTED
        End If
    Next i
    
    Part_Change user_select:=True

End Sub

Private Sub FeatureListRefresh_Click()
    Dim j As Integer
    For j = 0 To UBound(listParts(2, Me.PartListBox.ListIndex), 2)
        If listParts(2, Me.PartListBox.ListIndex)(1, j) = DataSources.ITEM_UNSELECTED Then
            listParts(2, Me.PartListBox.ListIndex)(1, j) = DataSources.ITEM_SELECTED
        End If
    Next j
    
    Feat_Change user_select:=True
    
End Sub

Private Sub FieldListRefresh_Click()
    Dim k As Integer
    For k = 0 To UBound(listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex), 2)
        If listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, k) = DataSources.ITEM_UNSELECTED Then
            listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, k) = DataSources.ITEM_SELECTED
        End If
    Next k

    CF_Change user_select:=True

End Sub


Private Sub PartListDelete_Click()
    Dim i As Integer, j As Integer, k As Integer
    For i = 0 To UBound(listParts, 2)
        If listParts(1, i) = DataSources.ITEM_SELECTED Then
            listParts(1, i) = DataSources.ITEM_UNSELECTED
        End If
        
        'Unselect all features and custom fields, even if the part is unselected originally
        For j = 0 To UBound(listParts(2, i), 2)
            If listParts(2, i)(1, j) = DataSources.ITEM_SELECTED Then
                listParts(2, i)(1, j) = DataSources.ITEM_UNSELECTED
                
                For k = 0 To UBound(listParts(2, i)(2, j), 2)
                    If listParts(2, i)(2, j)(1, k) = DataSources.ITEM_SELECTED Then
                        listParts(2, i)(2, j)(1, k) = DataSources.ITEM_UNSELECTED
                    End If
                Next k
            End If
        Next j
    Next i
    
    Part_Change user_select:=True
    
    
End Sub

Private Sub FeatureListDelete_Click()
    Dim j As Integer, k As Integer
    For j = 0 To UBound(listParts(2, Me.PartListBox.ListIndex), 2)
        If listParts(2, Me.PartListBox.ListIndex)(1, j) = DataSources.ITEM_SELECTED Then
            listParts(2, Me.PartListBox.ListIndex)(1, j) = DataSources.ITEM_UNSELECTED
        End If
        
        'Unselect all custom fields of all features in the part, Even if the feature was unselected originally
        For k = 0 To UBound(listParts(2, Me.PartListBox.ListIndex)(2, j), 2)
            If listParts(2, Me.PartListBox.ListIndex)(2, j)(1, k) = DataSources.ITEM_SELECTED Then
                listParts(2, Me.PartListBox.ListIndex)(2, j)(1, k) = DataSources.ITEM_UNSELECTED
            End If
        Next k
        
    Next j
    
    Feat_Change user_select:=False

End Sub

Private Sub FieldListDelete_Click()
    Dim k As Integer
    For k = 0 To UBound(listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex), 2)
        If listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, k) = DataSources.ITEM_SELECTED Then
            listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, k) = DataSources.ITEM_UNSELECTED
        End If
    Next k

    CF_Change user_select:=False


End Sub


Private Sub CopyFields_Button_Click()
    'If its a feature not applicable to being copied...
    Dim feat_selection As String: feat_selection = listParts(2, Me.PartListBox.ListIndex)(1, Me.FeatureListBox.ListIndex)
    If feat_selection = DataSources.ITEM_NOT_APPLICABLE Or feat_selection = DataSources.ITEM_NOT_FOUND Then
        MsgBox "Feature must be either Selected or Unselected," & vbCrLf & "This Feature cant be copied", vbExclamation
        Exit Sub
    End If
    
    
    Dim copyFields(6) As Variant
    Dim k As Integer, cf_val As String
    
    'Create a copy of the custom fields on the current feature
    For k = 0 To UBound(copyFields)
        cf_val = listParts(2, Me.PartListBox.ListIndex)(2, Me.FeatureListBox.ListIndex)(1, k)
            'NA features should be reflected as unselected
        If cf_val = DataSources.ITEM_NOT_FOUND Or cf_val = DataSources.ITEM_NOT_APPLICABLE Then
            copyFields(k) = DataSources.ITEM_UNSELECTED
        Else
            copyFields(k) = cf_val
        End If
    Next k
    
    Dim j As Integer
    For j = 0 To UBound(listParts(2, Me.PartListBox.ListIndex), 2)
        If j = Me.FeatureListBox.ListIndex Then GoTo skipFeat
        
        For k = 0 To UBound(copyFields)
            cf_val = listParts(2, Me.PartListBox.ListIndex)(2, j)(1, k)
            If cf_val = DataSources.ITEM_SELECTED Or cf_val = DataSources.ITEM_UNSELECTED Then
                listParts(2, Me.PartListBox.ListIndex)(2, j)(1, k) = copyFields(k)
            End If
        Next k
                
skipFeat:
    
    Next j
    
    CF_Change user_select:=False
        
    'The above logically switches the currently selected feature only,
        'If we want the other features to reflect the same status as this one,
        'we need to iterate through the features again and change their status
        
    'After re-evaluating the Feature selected, all other features in the part should have the same
    
    Dim feat_val As String, mapped_feat_val As String
    mapped_feat_val = listParts(2, Me.PartListBox.ListIndex)(1, Me.FeatureListBox.ListIndex)
    
    For j = 0 To UBound(listParts(2, Me.PartListBox.ListIndex), 2)
        If j = Me.FeatureListBox.ListIndex Then GoTo skipFeatAgain
            feat_val = listParts(2, Me.PartListBox.ListIndex)(1, j)
            If feat_val = DataSources.ITEM_NOT_FOUND Or feat_val = DataSources.ITEM_NOT_APPLICABLE Then
                GoTo skipFeatAgain
            Else
                listParts(2, Me.PartListBox.ListIndex)(1, j) = mapped_feat_val
            End If
skipFeatAgain:
    Next j
    
    ResetLists
    
    MsgBox "Custom Field Mapping Copied for the Selected Part", vbInformation
    
End Sub


Private Sub CopyFeatures_Button_Click()
    Dim i As Integer, j As Integer, k As Integer, q As Integer

    'If the part is not selected, then there's nothing to do
    If listParts(1, Me.PartListBox.ListIndex) <> DataSources.ITEM_SELECTED Then
        MsgBox "The Current Part is not selected," & vbCrLf & "the mapping cannot be copied over", vbExclamation
        Exit Sub
    End If
    
    'Check the other part numbers. If nothing else is applicable to selection, nothing will happen anyway
    Dim possible_changes As Integer
    For i = 0 To UBound(listParts, 2)
        If i = Me.PartListBox.ListIndex Then GoTo nextpart
        
        If listParts(1, i) = DataSources.ITEM_SELECTED Or listParts(1, i) = DataSources.ITEM_UNSELECTED Then GoTo partsOk
nextpart:
    Next i
        
    MsgBox "The Mapping would not effect any other parts anyway." & vbCrLf & "Aborting Operation", vbExclamation
    Exit Sub
partsOk:

    'Build the list of what we will copy to all of the other parts
    Dim copyList() As Variant
    ReDim Preserve copyList(0 To 2, UBound(listParts(2, Me.PartListBox.ListIndex), 2))
    
    For j = 0 To UBound(listParts(2, Me.PartListBox.ListIndex), 2)
        copyList(0, j) = listParts(2, Me.PartListBox.ListIndex)(0, j) 'Copy name
            'If a feature is NA, treat it as being unselected instead...
        If listParts(2, Me.PartListBox.ListIndex)(1, j) = DataSources.ITEM_NOT_FOUND Or listParts(2, Me.PartListBox.ListIndex)(1, j) = DataSources.ITEM_NOT_APPLICABLE Then
            copyList(1, j) = DataSources.ITEM_UNSELECTED
        Else
            copyList(1, j) = listParts(2, Me.PartListBox.ListIndex)(1, j)
        End If
        
        Dim fieldArr(6) As Variant
        For k = 0 To UBound(fieldArr)
                'IF custom field if NA, instead treat it as being unselected
            If listParts(2, Me.PartListBox.ListIndex)(2, j)(1, k) = DataSources.ITEM_NOT_APPLICABLE Or listParts(2, Me.PartListBox.ListIndex)(2, j)(1, k) = DataSources.ITEM_NOT_FOUND Then
                fieldArr(k) = DataSources.ITEM_UNSELECTED
            Else
                fieldArr(k) = listParts(2, Me.PartListBox.ListIndex)(2, j)(1, k)
            End If
        Next k
        
        'Set the customField Array and reset
        copyList(2, j) = fieldArr
        
        Erase fieldArr
    Next j
    
    'Start Copying the mapping
    For i = 0 To UBound(listParts, 2)
        If i = Me.PartListBox.ListIndex Then GoTo skipPartMap
        If listParts(1, i) = DataSources.ITEM_NOT_APPLICABLE Or listParts(1, i) = DataSources.ITEM_NOT_FOUND Then GoTo skipPartMap
        
        'Set the selection status to be the same as our current part's
        listParts(1, i) = listParts(1, Me.PartListBox.ListIndex)
        
            'For each feature in the part
        For j = 0 To UBound(listParts(2, i), 2)
            If listParts(2, i)(1, j) = DataSources.ITEM_NOT_FOUND Or listParts(2, i)(1, j) = DataSources.ITEM_NOT_APPLICABLE Then GoTo skipFeatMap
            
                'For each feature in our Copied List
            For k = 0 To UBound(copyList, 2)
                If copyList(0, k) = listParts(2, i)(0, j) Then  'If the names are the same
                
                    'Set the feature selection statuses to be the same
                    listParts(2, i)(1, j) = copyList(1, k)
                    
                    For q = 0 To UBound(copyList(2, k))
                        'If the custom field CAN be changed....
                        If listParts(2, i)(2, j)(1, q) = DataSources.ITEM_SELECTED Or listParts(2, i)(2, j)(1, q) = DataSources.ITEM_UNSELECTED Then
                            listParts(2, i)(2, j)(1, q) = copyList(2, k)(q)
                        End If
                    Next q
                End If
            Next k
        
skipFeatMap:
        Next j
skipPartMap:
    Next i
    
    ResetLists
    
    MsgBox "This Part's Feature/Field mappings have been" & vbCrLf & "Copied over to the other Parts", vbInformation
    
End Sub





'************************************************************
'*******************   Terminate   **************************
'************************************************************



Private Sub InsertSelectionButton_Click()
    'Manipulate the JSON object stored by removing unselected or not applicable items
    
    'Remove Parts as necessary
resetPartCheck:
    If json_routine_map.Count = 0 Then
        MsgBox "Didnt find any valid parts to send to the MeasurLink database", vbInformation
        GoTo quit_form
    End If
    
    Dim part_ind As Integer, feat_ind As Integer, cf_ind As Integer, i As Integer, j As Integer, k As Integer
    
    For part_ind = 1 To json_routine_map.Count
        For i = 0 To UBound(listParts, 2)
            If json_routine_map(part_ind)("name") = listParts(0, i) Then
                If listParts(1, i) <> DataSources.ITEM_SELECTED Then
                    json_routine_map.Remove part_ind
                    GoTo resetPartCheck
                Else  '**we found a match and the Part is selected, begin Feature validation
resetFeatCheck:
                        'If the part doesnt have features, delete the
                    If json_routine_map(part_ind)("features").Count = 0 Then
                        json_routine_map.Remove part_ind
                        GoTo resetPartCheck
                    End If
                    
                    For feat_ind = 1 To json_routine_map(part_ind)("features").Count
                        For j = 0 To UBound(listParts(2, i), 2)
                            If json_routine_map(part_ind)("features")(feat_ind)("name") = listParts(2, i)(0, j) Then  'found a name match
                                If listParts(2, i)(1, j) <> DataSources.ITEM_SELECTED Then
                                    json_routine_map(part_ind)("features").Remove feat_ind
                                    GoTo resetFeatCheck
                                Else  '**we found a match and the Feature is selected, begin Custom Field validation
                                    
resetFieldCheck:
                                    If json_routine_map(part_ind)("features")(feat_ind)("custom_fields").Count = 0 Then
                                        json_routine_map(part_ind)("features").Remove feat_ind
                                        GoTo resetFeatCheck
                                    End If
                                    
                                    For cf_ind = 1 To json_routine_map(part_ind)("features")(feat_ind)("custom_fields").Count
                                        For k = 0 To UBound(listParts(2, i)(2, j), 2)
                                            If json_routine_map(part_ind)("features")(feat_ind)("custom_fields")(cf_ind)("customFieldId") = listParts(2, i)(2, j)(0, k) Then
                                                If listParts(2, i)(2, j)(1, k) <> DataSources.ITEM_SELECTED Then
                                                    json_routine_map(part_ind)("features")(feat_ind)("custom_fields").Remove cf_ind
                                                    GoTo resetFieldCheck
                                                Else
                                                    GoTo nextfield  'we found the customField and its all good, skip over removal
                                                End If
                                            End If
                                        Next k
                                        
                                        'Got here without finding a match.... somehow
                                        json_routine_map(part_ind)("features")(feat_ind)("custom_fields").Remove cf_ind
                                        GoTo resetFieldCheck
                                    
nextfield:
                                    Next cf_ind
                                    
                                    GoTo nextfeat  'End with going to the next feat if its all good, skip over removal
                                End If
                            End If
                        Next j
                        
                        'We got here and didnt find a FeatureName to match against
                        json_routine_map(part_ind)("features").Remove feat_ind
                        GoTo resetFeatCheck
                        
nextfeat:
                    Next feat_ind
                    GoTo nextpart 'end with skipping over the whole, no partname check
                    
                End If
            End If
        Next i
        
        
        'We didnt find Part name match, theres a part num in the wb that isnt in the db
        json_routine_map.Remove part_ind
        GoTo resetPartCheck
        
nextpart:
    Next part_ind
    
    Dim dupe_json As Object
    Set dupe_json = json_routine_map  'make a backup before it becomes unloaded
    
    Unload Me
    
    ParseArray_ForUpload json_parts_info:=dupe_json
    
    Exit Sub
        
    
'----------Debugging
'    Dim clip As DataObject
'
'    Set clip = New DataObject
'    clip.SetText (JsonConverter.ConvertToJson(json_routine_map))
'    clip.PutInClipboard

'    MsgBox "stop here"
    
quit_form:
    Unload Me
    
    
End Sub





Private Sub UserForm_Click()

End Sub
