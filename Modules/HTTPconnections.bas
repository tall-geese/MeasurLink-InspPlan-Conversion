Attribute VB_Name = "HTTPconnections"
'*************************************************************
'*************************************************************
'*                  HTTPconnections
'*
'*  Connect to the API at Jade76 and
'*    SELECT, UPDATE or INSERT Custom Field Information
'*    for each Part's Feautres
'*
'*
'*
'*     'Dictionaries in python are translated as VBA Dictionaries, Lists like Collections, so.....
'*         '{'hello': 'world'}    ->   parsed("hello") ->> "world"
'*         '{'hello': {'goodbye': 'world'}}   -> parsed("hello")("goodbye") ->> "world"
'*         '{'hello': [{'goodbye': 'world'}]}   -> parsed("hello")(1)("goobye") ->> "world"
'*
'*         'Just keep in mind that Collections are 1 based.
'*         'When trying to flatten results, use TypeName() -> 'Collection' | 'Dictionary' | [some scalar]
'*             'To figure out how to iterate through the items
'*
'*************************************************************
'*************************************************************



'****************************************************
'**************   Main Routine   ********************
'****************************************************

Public Function send_http(url As String, method As String, payload As String, Optional q_params As Variant, Optional api_key As Variant) As String
    On Error GoTo HTTP_Err:

    Dim req As ServerXMLHTTP60
    Dim parsed As Object

    Set req = New ServerXMLHTTP60
    
    If Not IsMissing(q_params) Then
        'Set up the Url to add query parameters to the end
        'q_params(i)(0) -> key
        'q_params(i)(1) -> val
        
        url = url & "?"
        Dim i As Integer
        For i = 0 To UBound(q_params)
            url = url & q_params(i)(0) & "=" & q_params(i)(1)
        Next i
    End If
    
    With req
        'Set request headers here...
        .Open method, url, False   'We can do this asyncronously??
        .setRequestHeader "Content-Type", "application/json;charset=utf-8"
        .setRequestHeader "Accept", "application/json;charset=utf-8"
        
        If Not IsMissing(api_key) Then
            .setRequestHeader "X-Request-ID", api_key
            .setRequestHeader "Authorization", Environ("Username")
        End If
        
        .send payload
    End With

    Dim resp As String, header As String, headers As String
    headers = req.getAllResponseHeaders()
    
    Debug.Print (headers)
    Debug.Print (req.Status & vbTab & req.statusText)
    
    If req.Status <> 200 Then GoTo HTTP_Err
    'Should read the response type here and possible raise and error based on the different response types we can get
    
    send_http = req.responseText
    
    Exit Function
    
HTTP_Err:
    If req.readyState < 4 Then
        Err.Raise Number:=vbObjectError + 6010, Description:="send_http Error" & vbCrLf & vbCrLf & "No response from the server. The server may be down or the API service may not be running"

    ElseIf req.Status = 406 Or req.Status = 400 Or req.Status = 404 Then
        'Adding a user: Either not in QA department or they have already been reigstered
        Err.Raise Number:=vbObjectError + 6000 + req.Status, Description:=req.responseText
    Else
        'Unhandled HTTP Errors, Likely for Internal Server 500
        Err.Raise Number:=vbObjectError + 6000, Description:="send_http Error" & vbCrLf & headers & vbCrLf & "Status:" & req.Status & vbTab & req.statusText _
            & vbCrLf & "RequestBody: " & vbCrLf & payload & vbCrLf & vbclrf
    End If
End Function


'****************************************************
'************   Public Callables   ******************
'****************************************************

Public Function GetPartsInfo(part_numbers() As String) As Object
    'part_numbers() -> Part Numbers must be of Format "Part_Rev" like...
        '1906751_A

    On Error GoTo GetPartsInfo_Err:
    
    Dim payload As String
    payload = JsonConverter.ConvertToJson(part_numbers)
    
    Dim resp As String
    resp = send_http(url:=DataSources.JPMCML_PARTS_INFO, method:=DataSources.HTTP_POST, payload:=payload)
    
    Set GetPartsInfo = JsonConverter.ParseJson(resp)
    Exit Function
    
GetPartsInfo_Err:
    If Err.Number = vbObjectError + 6000 Then  'Unhandled Exceptions Like Internal Server Error
        MsgBox Err.Description
    ElseIf Err.Number = vbObjectError + 6010 Then  'Server Not Responding
        MsgBox Err.Description, vbExclamation
    Else
        MsgBox "Unexpected Exception Occured Func: HTTPConnections.GetPartsInfo() when parsing JSON" & vbCrLf & vbCrLf & Err.Description, vbCritical
    End If
End Function


Public Sub AddCurrentUser()
    On Error GoTo AddUser_Err:
    
    Dim returnMsg As Object
        
    Dim q_params(0) As Variant
    Dim q_name() As Variant
    q_name = Array("user", Environ("Username"))
    q_params(0) = q_name
    
    Dim resp As String
    resp = send_http(url:=DataSources.JPMCML_ADD_USER, method:=DataSources.HTTP_POST, payload:="", q_params:=q_params)
    
    Set returnMsg = JsonConverter.ParseJson(resp)
    
    MsgBox "User: " & Environ("Username") & " has successfully been added" & vbCrLf & vbCrLf & "An email should be arriving soon with the API Key file and Directions on where it should be stored" _
                & vbCrLf & vbCrLf _
                & "First Name: " & vbTab & vbTab & returnMsg("firstName") & vbCrLf _
                & "Last Name: " & vbTab & vbTab & returnMsg("lastName") & vbCrLf _
                & "Abbreviation: " & vbTab & vbTab & returnMsg("abbrev") & vbCrLf _
                & "Email: " & vbTab & vbTab & returnMsg("email")
    
    Exit Sub
    
AddUser_Err:
    If Err.Number = vbObjectError + 6000 Then  'Unhandled Exceptions like Internal Server Error
        MsgBox Err.Description, vbCritical
    ElseIf Err.Number = vbObjectError + 6010 Then   'Server /API service is down
        MsgBox Err.Description, vbCritical
    ElseIf Err.Number = vbObjectError + 6100 Then   'HTTP Repsonse, User Not allowed to be added.
        Set returnMsg = JsonConverter.ParseJson(Err.Description)
        
        MsgBox "User: " & Environ("Username") & " could not be registered, see the response below" & vbCrLf & vbCrLf & returnMsg("detail"), vbCritical
    Else
        MsgBox "Unexpected Exception Occured Func: HTTPConnections.AddCurrentUser()" & vbCrLf & vbCrLf & Err.Description, vbCritical
    End If
End Sub


Public Function AddCustomFields(payload As String, api_key As String) As String
    
    On Error GoTo addCustomFieldsErr:

    Dim resp As String
    resp = send_http(url:=DataSources.JPMCML_FIELDS_ADD, method:=DataSources.HTTP_POST, payload:=payload, api_key:=api_key)

    AddCustomFields = resp
    
    Exit Function

addCustomFieldsErr:
        
    If Err.Number = vbObjectError + 6010 Or Err.Number = vbObjectError + 6404 Then 'Server Down / Part,Feature Combo not found
        MsgBox Err.Description, vbExclamation
        
    ElseIf Err.Number = vbObjectError + 6400 Or Err.Number = vbObjectError + 6000 Then  'Custom Field already exists, Internal Server Error
        MsgBox Err.Description, vbCritical
        
    Else   'Unhandle Exceptions
        MsgBox "Unexpected Exception Occured Func: HTTPConnections.AddCustomFields()" & vbCrLf & vbCrLf & Err.Description, vbCritical
    End If
End Function

Public Function UpdateCustomFields(payload As String, api_key As String) As String
    
    On Error GoTo updateCustomFieldsErr:

    Dim resp As String
    resp = send_http(url:=DataSources.JPMCML_FIELDS_UPDATE, method:=DataSources.HTTP_PUT, payload:=payload, api_key:=api_key)

    UpdateCustomFields = resp
        
    Exit Function

updateCustomFieldsErr:
        
    If Err.Number = vbObjectError + 6010 Or Err.Number = vbObjectError + 6404 Then 'Server Down / Part,Feature Combo not found
        MsgBox Err.Description, vbExclamation
        
    ElseIf Err.Number = vbObjectError + 6400 Or Err.Number = vbObjectError + 6000 Then  'User not allowed, Internal Server Error
        MsgBox Err.Description, vbCritical
        
    Else   'Unhandle Exceptions
        MsgBox "Unexpected Exception Occured Func: HTTPConnections.UpdateCustomFields()" & vbCrLf & vbCrLf & Err.Description, vbCritical
    End If
End Function













''''''''''''''''Test





Sub test_our_parts()
   Dim parts(1) As String
   parts(0) = "9999999_A"
   parts(1) = "1906752_A"
   'go back and test what happes if we send a bad part number
   
   Dim out_arr As Object
   Set out_arr = GetPartsInfo(parts)
   
   Dim a As DataObject
   Set a = New DataObject
   a.SetText JsonConverter.ConvertToJson(out_arr)
   a.PutInClipboard
   
End Sub








'''''OLD

Public Function GetCustomFieldInfo(partNumbers() As String) As Variant()
    'pass in the url here to the main calling function
    'parse out the respose for what the user expects
    Dim parsed As Object
    Dim StartTime As Double
    StartTime = Timer
    
    Set parsed = JsonConverter.ParseJson(req.responseText)
    Debug.Print (parsed("hello")(1)("goodbye"))
    
    Debug.Print ("Completed in: " & Timer - StartTime & " seconds")

End Function



Private Sub test()
    Dim a As String
    a = "asdfads"
    test2 something:=a
End Sub

Private Sub test2(Optional something As Variant)
    Debug.Print IsMissing(something)

End Sub







