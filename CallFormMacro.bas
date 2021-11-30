Attribute VB_Name = "CallFormMacro"
Sub Architectural_Family(RibbonControl As IRibbonControl)
ThisWorkbook.Sheets("A").Unprotect Password:="ttdg"

Worksheets("A").Activate
UserFormA.Show
End Sub
Sub Electrical_Family(RibbonControl As IRibbonControl)
ThisWorkbook.Sheets("E").Unprotect Password:="ttdg"

Worksheets("E").Activate
UserFormE.Show
End Sub
Sub FAFF_Family(RibbonControl As IRibbonControl)
ThisWorkbook.Sheets("F").Unprotect Password:="ttdg"

Worksheets("F").Activate
UserFormF.Show
End Sub
Sub Mechanical_Family(RibbonControl As IRibbonControl)
ThisWorkbook.Sheets("M").Unprotect Password:="ttdg"

Worksheets("M").Activate
UserFormM.Show
End Sub
Sub Plumbing_Family(RibbonControl As IRibbonControl)
ThisWorkbook.Sheets("P").Unprotect Password:="ttdg"

Worksheets("P").Activate
UserFormP.Show
End Sub
Sub Structural_Family(RibbonControl As IRibbonControl)
ThisWorkbook.Sheets("S").Unprotect Password:="ttdg"

Worksheets("S").Activate
UserFormS.Show
End Sub

