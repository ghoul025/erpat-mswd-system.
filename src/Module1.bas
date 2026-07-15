Attribute VB_Name = "Module1"
Public Sub PrintFrontAndBack()

    Dim arrSheets As Variant
    arrSheets = Array("PRINT_FRONT", "PRINT_BACK")

    ThisWorkbook.Worksheets(arrSheets).Select

    ActiveWindow.SelectedSheets.PrintOut _
        Copies:=1, _
        Collate:=True

    ' Return to main sheet (optional)
    ThisWorkbook.Worksheets(SRC_SHEET).Select

End Sub

