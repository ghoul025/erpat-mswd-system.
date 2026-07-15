VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    With ThisWorkbook.Worksheets("Sheet1")
        ' 1. Temporarily unprotect to modify cell properties safely
        .Unprotect Password:="erpatsystem"
        
        ' 2. Force all cells to be locked so contents cannot be edited
        .Cells.Locked = True
        
        ' 3. Apply protection and allow selection rules
        .Protect Password:="erpatsystem", _
                 DrawingObjects:=True, _
                 Contents:=True, _
                 Scenarios:=True, _
                 UserInterfaceOnly:=True, _
                 AllowFiltering:=True, _
                 AllowSorting:=True, _
                 AllowUsingPivotTables:=True
                 
        ' 4. Allow users to physically click/select all locked cells
        .EnableSelection = xlNoRestrictions
        
        ' 5. Lock table layout to prevent automatic column resizing
        ' Change "YourTableName" to the actual name of your Excel table
        ' .ListObjects("tblResidents").MaintainColumnWidths = True
    End With

End Sub


