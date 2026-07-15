Attribute VB_Name = "CLEAR_MODULE"
Sub ClearAllSlicers()
    Dim sc As SlicerCache
    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc
End Sub
