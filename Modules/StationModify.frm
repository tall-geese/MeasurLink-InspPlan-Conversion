VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StationModify 
   Caption         =   "Modify Available Cells and Stations"
   ClientHeight    =   6885
   ClientLeft      =   -255
   ClientTop       =   -930
   ClientWidth     =   11145
   OleObjectBlob   =   "StationModify.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StationModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private json_config As Object
Private listCells() As Variant
    'listCells(0, i)  <-- Cell Name
    'listCells(1, i)  <-- Cell Color rgb(xxx, xxx, xxx)
    'listCells(2, i)  <-- Cell Station List
        'listCells(2, i)(0, j)  <-- Station Name
        'listCells(2, i)(1, j)  <-- Resource Name

Private station_text_boxes_interacted As Boolean
Private cell_text_boxes_interacted As Boolean

Private Const CELL_TEXT_DEFAULT = "Enter New Cell Name"
Private Const STATION_TEXT_DEFAULT = "Enter Station Name"
Private Const RESOURCE_TEXT_DEFAULT = "Enter Resource Name"


'*************************************************************
'*************************************************************
'*                  Public Callbacks
'*************************************************************
'*************************************************************

Public Sub Unravel_cells(json As Object)
    Set json_config = json
    Erase listCells
    cell_text_boxes_interacted = False
    station_text_boxes_interacted = False
    
    For Each cell In json
        Dim cell_upper As Integer
        If (Not listCells) = -1 Then
            ReDim Preserve listCells(2, 0)
            listCells(0, 0) = cell("name")
            listCells(1, 0) = cell("color")
            cell_upper = 0
        Else
            cell_upper = UBound(listCells, 2) + 1
            ReDim Preserve listCells(2, cell_upper)
            listCells(0, cell_upper) = cell("name")
            listCells(1, cell_upper) = cell("color")
        
        End If
        
        Dim stations() As Variant
        For Each station In cell("stations")
            If (Not stations) = -1 Then
                ReDim Preserve stations(1, 0)
                stations(0, 0) = station("name")
                stations(1, 0) = station("resource")
            Else
                Dim stations_upper As Integer
                stations_upper = UBound(stations, 2) + 1
                ReDim Preserve stations(1, stations_upper)
                stations(0, stations_upper) = station("name")
                stations(1, stations_upper) = station("resource")
            End If
        Next station
        
        listCells(2, cell_upper) = stations
        Erase stations
    
    Next cell
    
    Call RefreshCells
    
End Sub

    'Public callback from the StationColors after it successfully closes
Public Sub Color_Received(RGB As String)
    'If we're changing an existing cell
    Dim ind As Integer
    ind = Me.Cell_listBox.ListIndex
    If ind <> -1 Then
        listCells(1, ind) = RGB
    End If
    
    
    'Otherwise, lets save it for when we are going to create a new cell
    ResetButtonColor color:=RGB
    RefreshCells
    
End Sub



'*************************************************************
'*************************************************************
'*                  Submit Button
'*************************************************************
'*************************************************************

Private Sub SubmitButton_Click()
    'Turn the array data back into Json
    Dim cell_coll As Collection
    Set cell_coll = New Collection: Dim i As Integer, j As Integer
    
    For i = 0 To UBound(listCells, 2)
        Set cell_dict = New Dictionary: Dim stations() As Variant: Set station_coll = New Collection
        cell_dict.Add "name", listCells(0, i)
        cell_dict("color") = listCells(1, i)
                
        stations = listCells(2, i)
        If (Not stations) = -1 Then GoTo next_cell
        
        For j = 0 To UBound(stations, 2)
            Set station_dict = New Dictionary
            station_dict.Add "name", listCells(2, i)(0, j)
            station_dict.Add "resource", listCells(2, i)(1, j)
        
            station_coll.Add station_dict
        Next j
next_cell:

        cell_dict.Add "stations", station_coll
        Erase stations
        cell_coll.Add cell_dict
    Next i
    
    If cell_coll.Count = 0 Then Exit Sub
    
    RibbonCommands.UpdateConfig json:=cell_coll
    Unload Me
    

End Sub



'*************************************************************
'*************************************************************
'*                  Cell Controls
'*************************************************************
'*************************************************************

Private Sub AddCell_Click()
    Dim cell_name As String
    cell_name = Me.TextBox_Cell.Value
    cell_name = Replace(cell_name, "\", "")
    
    Me.TextBox_Cell.Value = ""
    
    If cell_name = vbNullString Or cell_name = CELL_TEXT_DEFAULT Then
        MsgBox "Enter a Cell Name", vbInformation
        Exit Sub
    End If
    

    Dim color As Long
    color = Me.ChangeCellColors.ForeColor
    
    Dim R As String, G As String, B As String, RGB As String, hex_color As String
    hex_color = GetHexFromLong(Me.ChangeCellColors.BackColor)
    R = GetRGBFromHex(hex_color, "R")
    G = GetRGBFromHex(hex_color, "G")
    B = GetRGBFromHex(hex_color, "B")
    
    RGB = "rgb(" & R & "," & G & "," & B & ")"
    Dim empT() As Variant
    
    If (Not listCells) = -1 Then
        ReDim Preserve listCells(2, 0)
        listCells(0, 0) = cell_name
        listCells(1, 0) = RGB
        listCells(2, 0) = empT
    Else
        Dim upper As Integer
        upper = UBound(listCells, 2) + 1
        ReDim Preserve listCells(2, upper)
        listCells(0, upper) = cell_name
        listCells(1, upper) = RGB
        listCells(2, upper) = empT
    End If
    
    Call RefreshCells
    
    
End Sub

Private Sub Cells_EraseCellButton_Click()
    If Me.Cell_listBox.ListIndex = -1 Then Exit Sub
    
    listCells = DropRows(listCells, Me.Cell_listBox.ListIndex)
    RefreshCells
    RefreshStations
    
End Sub



Private Sub Cell_listBox_Change()
    If Me.Cell_listBox.ListIndex = -1 Then Exit Sub
    
    Me.Stations_listBox.Clear
    
    Dim stations() As Variant
    stations = listCells(2, Me.Cell_listBox.ListIndex)
    If (Not stations) = -1 Then GoTo color_change
    stations = Application.Transpose(stations)
    stations = Force2D(stations)
    
    Me.Stations_listBox.list = stations
    
color_change:
    Dim color As String
    color = listCells(1, Me.Cell_listBox.ListIndex)
    ResetButtonColor color:=color
    
End Sub



Private Sub ChangeCellColors_Click()
    Dim color As String
    Load StationColors
    
    If Not Me.ChangeCellColors.BackColor = DataSources.COLOR_DEFAULT Then
        color = listCells(1, Me.Cell_listBox.ListIndex)
        color = Replace(color, " ", "")
        
        Set re = New RegExp
        re.Global = False
        re.Pattern = "\((\d*),(\d*),(\d*)\)"
        
        Dim match As Object
        Set match = re.Execute(color)
        If Not match Is Nothing Then
            'send something over to the StationColors form
            StationColors.ColorSeed red:=match(0).SubMatches(0), green:=match(0).SubMatches(1), blue:=match(0).SubMatches(2)
        End If
        
    End If
    
    StationColors.Show vbModeless
    StationColors.Repaint
    
    
End Sub

Private Sub TextBox_Cell_Enter()
    If Not cell_text_boxes_interacted Then
        ClearCellBox
    End If
    
    Me.ChangeCellColors.BackColor = DataSources.COLOR_DEFAULT
    Me.Cell_listBox.ListIndex = -1
End Sub



'*************************************************************
'*************************************************************
'*                  Station Controls
'*************************************************************
'*************************************************************

Private Sub AddStation_Click()
    'If they didnt fill out information, then quit out
    Dim station As String, resource As String
    station = Me.TextBox_Station.Value
    resource = Me.TextBox_Resource.Value
    
    station = Replace(station, "\", "")
    resource = Replace(resource, "\", "")
    
    Me.TextBox_Station.Value = ""
    Me.TextBox_Resource.Value = ""
    
    If station = vbNullString Or resource = vbNullString Or resource = RESOURCE_TEXT_DEFAULT Or station = STATION_TEXT_DEFUALT Then
        MsgBox "You need to Enter Information", vbInformation
        Exit Sub
    End If
    
    If Me.Cell_listBox.ListIndex = -1 Then
        MsgBox "No Cell Selected", vbInformation
        Exit Sub
    End If
    
    Dim stations() As Variant, upper As Integer
    stations = listCells(2, Me.Cell_listBox.ListIndex)
    If (Not stations) = -1 Then
        ReDim Preserve stations(1, 0)
        stations(0, 0) = station
        stations(1, 0) = resource
    Else
        upper = UBound(stations, 2) + 1
        ReDim Preserve stations(1, upper)
        stations(0, upper) = station
        stations(1, upper) = resource
    End If
    
    listCells(2, Me.Cell_listBox.ListIndex) = stations
    
    RefreshStations
    
End Sub


Private Sub Stations_EraseStationButton_Click()
    If Me.Cell_listBox.ListIndex = -1 Or Me.Stations_listBox.ListIndex = -1 Then
        Exit Sub
    End If
    
    Dim stations() As Variant
    stations = listCells(2, Me.Cell_listBox.ListIndex)
    If (Not stations) = -1 Then Exit Sub
    
    stations = DropRows(stations, Me.Stations_listBox.ListIndex)
    listCells(2, Me.Cell_listBox.ListIndex) = stations
    
    RefreshStations

End Sub


Private Sub TextBox_Resource_Enter()
    If Not station_text_boxes_interacted Then
        ClearStationBoxes
    End If
End Sub


Private Sub TextBox_Station_Enter()
    If Not station_text_boxes_interacted Then
        ClearStationBoxes
    End If
End Sub


'*************************************************************
'*************************************************************
'*                  Helper Functions
'*************************************************************
'*************************************************************

Private Sub RefreshCells()
    If (Not listCells) = -1 Then
        Me.Cell_listBox.Clear
        Exit Sub
    End If

    Dim list_box() As Variant, cols() As Variant
    cols = Array(0, 1)
    list_box = SliceColumns(listCells, cols)
    list_box = Application.Transpose(list_box)
    list_box = Force2D(list_box)
    
    Me.Cell_listBox.list = list_box

End Sub

Private Sub RefreshStations()
    Me.Stations_listBox.Clear
    Dim stations() As Variant, cols() As Variant
    If Me.Cell_listBox.ListIndex = -1 Then Exit Sub
    
    stations = listCells(2, Me.Cell_listBox.ListIndex)
    If (Not stations) = -1 Then Exit Sub
    
    stations = Application.Transpose(stations)
    stations = Force2D(stations)
    
    Me.Stations_listBox.list = stations
End Sub



Public Sub ResetButtonColor(color As String)
    color = Replace(color, " ", "")
    
    Set re = New RegExp
    re.Global = False
    re.Pattern = "\((\d*),(\d*),(\d*)\)"
    
    Dim match As Object
    Set match = re.Execute(color)
    If Not match Is Nothing Then
        Me.ChangeCellColors.BackColor = RGB(CInt(match(0).SubMatches(0)), CInt(match(0).SubMatches(1)), match(0).SubMatches(2))
    Else
        Me.ChangeCellColors.BackColor = DataSources.COLOR_DEFAULT
    End If

End Sub


Public Sub ClearCellBox()
    Me.TextBox_Cell.Value = ""
    cell_text_boxes_interacted = True
End Sub

Public Sub ClearStationBoxes()
    Me.TextBox_Resource.Value = ""
    Me.TextBox_Station.Value = ""
    station_text_boxes_interacted = True
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

Public Function DropRows(arr() As Variant, ind As Integer) As Variant()
    'Given and array, arr(x, y), drops the i'th row from y dimension. Based 0
    
    Dim cols As Integer, rows As Integer, out() As Variant
    cols = UBound(arr)
    rows = UBound(arr, 2)
    
    If rows = 0 Then
        DropRows = out
        Exit Function
    End If
    
    ReDim Preserve out(cols, rows - 1)
    
    Dim i As Integer, j As Integer, offset As Integer
    For i = 0 To cols
        For j = 0 To rows - 1
            If j = ind Then offset = 1  'When we hit the dropRow, start skipping to the next
            out(i, j) = arr(i, j + offset)
        Next j
        offset = 0
    Next i
    
    DropRows = out
    
End Function


Function GetHexFromLong(longColor As Long) As String

    Dim R As String
    Dim G As String
    Dim B As String
    
    R = Format(Application.WorksheetFunction.Dec2Hex(longColor Mod 256), "00")
    G = Format(Application.WorksheetFunction.Dec2Hex((longColor \ 256) Mod 256), "00")
    B = Format(Application.WorksheetFunction.Dec2Hex((longColor \ 65536) Mod 256), "00")
    
    GetHexFromLong = "#" & R & G & B

End Function

Function GetRGBFromHex(hexColor As String, RGB As String) As String

    hexColor = VBA.Replace(hexColor, "#", "")
    hexColor = VBA.Right$("000000" & hexColor, 6)
    
    Select Case RGB
        Case "B"
            GetRGBFromHex = VBA.val("&H" & VBA.Mid(hexColor, 5, 2))
        Case "G"
            GetRGBFromHex = VBA.val("&H" & VBA.Mid(hexColor, 3, 2))
        Case "R"
            GetRGBFromHex = VBA.val("&H" & VBA.Mid(hexColor, 1, 2))
    End Select
End Function

Private Sub UserForm_Click()

End Sub
