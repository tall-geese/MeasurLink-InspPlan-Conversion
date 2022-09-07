VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CrystalReport_Form 
   Caption         =   "Create Crystal Reports"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9450.001
   OleObjectBlob   =   "CrystalReport_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CrystalReport_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
'*************************************************************
'*                  CrystalReport_Form
'*
'*  Create the request to send to the server for
'*      Generating the Crystal Reports Needed
'*************************************************************
'*************************************************************




Private Sub Parts_ClearSingle_Click()
    If Me.PartNumbers_ListBox.ListIndex <> -1 Then
        Me.PartNumbers_ListBox.RemoveItem (Me.PartNumbers_ListBox.ListIndex)
    End If
End Sub

Private Sub Parts_ClearAll_Click()
    Me.PartNumbers_ListBox.Clear
End Sub

Private Sub Parts_SelectMulti_Click()
    Dim parts() As String, i As Integer, rev As String
    parts = Sheets("Variables").GetPartNumberOrNumbers()
    
    If (Not parts) = -1 Then Exit Sub
    
    rev = Sheets("START HERE").GetRevision()
    If rev = vbNullString Then
        MsgBox "Revision Not Found", vbCritical
        Exit Sub
    End If
    
    For i = 0 To UBound(parts)
        parts(i) = parts(i) & "_" & rev
    Next i
    
    Me.PartNumbers_ListBox.Clear
    Me.PartNumbers_ListBox.list = parts

End Sub

Private Sub Submit_Button_Click()
    If UBound(Me.PartNumbers_ListBox.list) = -1 Then Me.Hide
    
    
    'Build the optional Parameter Queries for the API call,
        'Only Customer is required
    Dim params() As Variant, coll As Collection, customer As String
    Set coll = New Collection
    coll.Add Me.Option_FA_ALL
    coll.Add Me.Option_FA_FIRST
    coll.Add Me.Option_FI_ALL
    coll.Add Me.Option_IP_ALL
    
    customer = Me.Customer_TextBox.Value
    ReDim Preserve params(0)
    Dim cust(1) As Variant
    cust(0) = "cust"
    cust(1) = customer
    params(0) = cust
    
    For Each member In coll
        If member.Value = True Then
            ReDim Preserve params(UBound(params) + 1)
            
            Dim skip(1) As Variant
            skip(0) = "skip"
            skip(1) = member.Tag
            params(UBound(params)) = skip
        End If
    Next member

    If Me.Option_Verify.Value = True Then
        ReDim Preserve params(UBound(params) + 1)
        Dim verify(1) As Variant
        verify(0) = "verify"
        verify(1) = True
        params(UBound(params)) = verify
    End If
    
    If Me.Option_Notify.Value = True Then
        ReDim Preserve params(UBound(params) + 1)
        Dim notify(1) As Variant
        notify(0) = "notify"
        notify(1) = True
        params(UBound(params)) = notify
    End If

    Dim parts() As Variant
    parts = Me.PartNumbers_ListBox.list
    parts = Application.Transpose(parts)

    RibbonCommands.Submit_Crystal_Reports customer:=customer, parts:=parts, params:=params
    
    Me.Hide

End Sub

