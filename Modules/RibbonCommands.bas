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

Public Sub LoadDataValidations(ByRef control As IRibbonControl)
    Dim valWb As Workbook
    Set valWb = Workbooks.Open(Filename:=DataSources.DATA_VALIDATION_PATH, UpdateLinks:=0, ReadOnly:=True)
    valWb.Sheets("Description").SetValReference (ThisWorkbook.Name)
    valWb.Sheets("StandardComments").SetValReference (ThisWorkbook.Name)
    valWb.Sheets("InspMethods").SetValReference (ThisWorkbook.Name)
End Sub

