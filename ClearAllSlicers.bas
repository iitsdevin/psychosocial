Attribute VB_Name = "ClearAllSlicers"
' Optional one-click "clear all slicers" for the Principal wellbeing email
' triage tool. Excel cannot clear all slicers without a macro, so this is
' provided separately for teams whose IT policy allows macros.
'
' To install:
'   1. Open the triage tool in Excel and press Alt+F11.
'   2. File > Import File... and choose this ClearAllSlicers.bas.
'   3. Save the workbook as "Excel Macro-Enabled Workbook (*.xlsm)".
'   4. (Optional) Insert > Shapes on the Interactive triage tab, draw a
'      button, right-click it > Assign Macro > ClearAllSlicers.
'
' Without macros, each slicer is cleared with the small funnel-with-x
' button in its top right corner.

Public Sub ClearAllSlicers()
    Dim sc As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc
End Sub
