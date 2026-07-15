VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const SHEET_PWD As String = "erpatsystem"

Private Sub Worksheet_Activate()

    ' 1. Optimize window view and hide the horizontal bar
    With ActiveWindow
        .Zoom = 83
        .DisplayHorizontalScrollBar = False
    End With

    ' 2. Force view to top-left corner on entry
    ' Application.GoTo Me.Range("A1"), True

    ' 3. Protect sheet interface safely
    Me.Protect Password:=SHEET_PWD, _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True, _
               AllowUsingPivotTables:=True
               
    ' 4. Ensure users can still select cells within limits
    Me.EnableSelection = xlNoRestrictions

    ' ==========================================================
    ' 5. DISABLE HORIZONTAL SCROLL PAST COLUMN P
    ' ==========================================================
    ' Restricts grid navigation strictly from Column A to Column P (down to row 1000)
    Me.ScrollArea = "A1:P100000"

End Sub

' Private Sub Worksheet_Change(ByVal Target As Range)

 '     Dim tbl As ListObject
  '    On Error Resume Next
   '   Set tbl = Me.ListObjects("tblResidents")
    '  On Error GoTo 0
'
 '     If tbl Is Nothing Then Exit Sub
  '    If tbl.DataBodyRange Is Nothing Then Exit Sub
'
 '     ' only fire when the changed cell is inside the table
  '    If Intersect(Target, tbl.DataBodyRange) Is Nothing Then Exit Sub
'
 '     ' small debounce — ignore multi-cell pastes of > 50 cells
  '    ' to avoid rebuilding mid-paste on large imports
   '   If Target.Cells.Count > 50 Then Exit Sub
'
 '     Application.EnableEvents = False
  '    BuildERPATDashboard          ' call the main builder
   '   Application.EnableEvents = True

 ' End Sub

