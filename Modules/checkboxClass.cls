VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "checkboxClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents cbEvents As MSForms.checkBox
Attribute cbEvents.VB_VarHelpID = -1
Public box As MSForms.checkBox
Public station_name As String

Private Sub cbEvents_Change()
    Dim empT() As Variant
    StationMapping.MapToRoutines station:=station_name, rt_inds:=empT, checkBox_Hook:=True, checkBox_Value:=box.Value
    StationMapping.EvalSelectionStatus checkBox_Hook:=True
End Sub

Private Sub cbEvents_Click()
End Sub
