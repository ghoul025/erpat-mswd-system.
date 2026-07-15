Attribute VB_Name = "Module3"
Sub InstantBulkExport()
    Dim comp As Object
    Dim exportPath As String
    Dim ext As String
    
    ' Automatically sets the export destination to an /src/ folder
    exportPath = ThisWorkbook.Path & "\src\"
    
    ' Creates the folder automatically if it isn't there
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    
    ' Loops through absolutely everything (Sheets, Modules, Forms) instantly
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1 ' Standard Module
                ext = ".bas"
            Case 2 ' Class Module
                ext = ".cls"
            Case 3 ' UserForm
                ext = ".frm"
            Case Else ' Sheet Codes (Dashboard_Engine, DASHBOARD, ThisWorkbook)
                ext = ".cls"
        End Select
        
        ' Exports the item instantly if it has lines of code written inside
        If comp.CodeModule.CountOfLines > 0 Then
            comp.Export exportPath & comp.Name & ext
        End If
    Next comp
    
    MsgBox "All modules, forms, and sheet codes exported to /src/ instantly!", vbInformation, "Success"
End Sub

