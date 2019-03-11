Attribute VB_Name = "Module_template"
Sub loadSummary(ByVal control As IRibbonControl)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    tmpltFromSmmry
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub tmpltFromSmmry()

  Dim listNames As Variant
    Set listNames = getSummaryNameList()
    For Each nm In listNames
        frm_Editor.cmb_Names.AddItem nm
    Next
    ThisWorkbook.Worksheets("Settings").Range("summary_destination") = 2
    frm_Editor.Show
    Unload frm_Editor

End Sub
