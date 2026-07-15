VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmResident 
   Caption         =   "REGISTRATION FORM"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19350
   OleObjectBlob   =   "frmResident.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmResident"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mEditingRowIndex As Long

Private Sub UserForm_Initialize()
    PinToTop
    LoadBarangays
    LoadCivilStatus
    SetupList
    StartNewRecord
    RefreshList
End Sub

Private Sub UserForm_Activate()
    PinToTop
End Sub

Private Sub PinToTop()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + 20
    Me.Top = Application.Top + 120
End Sub

Private Sub LoadBarangays()

    On Error GoTo fallback

    Dim rng As Range
    Dim c As Range

    Set rng = ThisWorkbook.Names(BRGY_RANGE).RefersToRange

    cboBarangay.Clear

    For Each c In rng.Cells
        If Trim(c.Value) <> "" Then cboBarangay.AddItem c.Value
    Next c

    Exit Sub

fallback:

    cboBarangay.Clear
    cboBarangay.Style = fmStyleDropDownCombo

End Sub

Private Sub LoadCivilStatus()

    cboCivilStatus.Clear
    cboCivilStatus.AddItem "SINGLE"
    cboCivilStatus.AddItem "MARRIED"
    cboCivilStatus.AddItem "WIDOWED"
    cboCivilStatus.AddItem "SEPARATED"

End Sub

Private Sub SetupList()

    lstRecords.Clear
    lstRecords.ColumnCount = 4
    lstRecords.ColumnWidths = "70;120;120;150"

End Sub

Private Sub RefreshList(Optional ByVal keyword As String = "")

    Dim tbl As ListObject
    Dim i As Long

    Dim idv As String
    Dim sur As String
    Dim fir As String
    Dim addr As String

    Dim k As String

    Set tbl = GetRegistryTable()
    lstRecords.Clear

    If tbl.ListRows.Count = 0 Then Exit Sub

    k = LCase(Trim(keyword))

    For i = 1 To tbl.ListRows.Count

        idv = CStr(tbl.DataBodyRange.Cells(i, 1).Value)
        sur = CStr(tbl.DataBodyRange.Cells(i, 2).Value)
        fir = CStr(tbl.DataBodyRange.Cells(i, 3).Value)
        addr = CStr(tbl.DataBodyRange.Cells(i, 8).Value)

        If k = "" Or InStr(1, LCase(idv & " " & sur & " " & fir & " " & addr), k) > 0 Then

            lstRecords.AddItem idv
            lstRecords.List(lstRecords.ListCount - 1, 1) = sur
            lstRecords.List(lstRecords.ListCount - 1, 2) = fir
            lstRecords.List(lstRecords.ListCount - 1, 3) = addr

        End If

    Next i

End Sub

Private Sub txtSearch_Change()
    RefreshList txtSearch.Value
End Sub

Private Sub lstRecords_Click()
    LoadSelectedRecord
End Sub

Private Sub LoadSelectedRecord()

    Dim tbl As ListObject
    Dim f As Range
    Dim i As Long
    Dim selID As String

    If lstRecords.ListIndex < 0 Then Exit Sub

    selID = CStr(lstRecords.List(lstRecords.ListIndex, 0))

    Set tbl = GetRegistryTable()

    Set f = tbl.ListColumns(1).DataBodyRange.Find(selID, , xlValues, xlWhole)

    If f Is Nothing Then Exit Sub

    i = f.Row - tbl.DataBodyRange.Row + 1
    mEditingRowIndex = i

    txtID.Value = tbl.DataBodyRange.Cells(i, 1).Value
    txtSurname.Value = tbl.DataBodyRange.Cells(i, 2).Value
    txtFirstname.Value = tbl.DataBodyRange.Cells(i, 3).Value
    txtMiddlename.Value = tbl.DataBodyRange.Cells(i, 4).Value

    txtBirthdate.Value = FormatCellDate(tbl.DataBodyRange.Cells(i, 5).Value)

    cboCivilStatus.Value = tbl.DataBodyRange.Cells(i, 7).Value
    cboBarangay.Value = tbl.DataBodyRange.Cells(i, 8).Value

    txtPContact.Value = tbl.DataBodyRange.Cells(i, 9).Value
    txtEmergency.Value = tbl.DataBodyRange.Cells(i, 10).Value
    txtContact.Value = tbl.DataBodyRange.Cells(i, 11).Value
    txtOccupation.Value = tbl.DataBodyRange.Cells(i, 12).Value

    txtDateIssued.Value = FormatCellDate(tbl.DataBodyRange.Cells(i, 13).Value)

    cmdSave.Enabled = False
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True

End Sub

Private Sub cmdPickBirth_Click()

    Dim d As Variant

    d = frmCalendar.PickDate(txtBirthdate.Value)

    If IsDate(d) Then
        txtBirthdate.Value = Format(d, "mm/dd/yyyy")
    End If

End Sub

Private Sub cmdPickIssued_Click()

    Dim d As Variant

    d = frmCalendar.PickDate(txtDateIssued.Value)

    If IsDate(d) Then
        txtDateIssued.Value = Format(d, "mm/dd/yyyy")
    End If

End Sub

Private Sub txtSurname_Change(): ForceUpper txtSurname: End Sub
Private Sub txtFirstname_Change(): ForceUpper txtFirstname: End Sub
Private Sub txtMiddlename_Change(): ForceUpper txtMiddlename: End Sub
Private Sub txtEmergency_Change(): ForceUpper txtEmergency: End Sub
Private Sub txtOccupation_Change(): ForceUpper txtOccupation: End Sub

Private Sub ForceUpper(ByVal tb As MSForms.TextBox)

    Dim p As Long

    p = tb.SelStart
    tb.Value = UCase(tb.Value)
    tb.SelStart = p

End Sub

Private Sub cmdNew_Click()
    StartNewRecord
End Sub

Private Sub cmdClear_Click()
    ClearFields True
End Sub

Private Sub StartNewRecord()

    mEditingRowIndex = 0

    txtID.Value = NextID()
    txtDateIssued.Value = Format(Date, "mm/dd/yyyy")

    ClearFields True

    cmdSave.Enabled = True
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False

End Sub

Private Sub ClearFields(ByVal KeepID As Boolean)

    If Not KeepID Then txtID.Value = ""

    txtSurname.Value = ""
    txtFirstname.Value = ""
    txtMiddlename.Value = ""
    txtBirthdate.Value = ""

    cboCivilStatus.Value = ""
    cboBarangay.Value = ""

    txtPContact.Value = ""
    txtEmergency.Value = ""
    txtContact.Value = ""

    txtOccupation.Value = ""

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Function ValidateInputs() As Boolean

    ValidateInputs = False

    If Trim(txtSurname.Value) = "" Or Trim(txtFirstname.Value) = "" Then
        MsgBox "Surname and Firstname are required.", vbExclamation
        Exit Function
    End If

    If Trim(cboCivilStatus.Value) = "" Then
        MsgBox "Civil Status is required.", vbExclamation
        Exit Function
    End If

    If Trim(cboBarangay.Value) = "" Then
        MsgBox "Address is required.", vbExclamation
        Exit Function
    End If

    Dim bd As Variant
    Dim di As Variant

    bd = ParseMDY(txtBirthdate.Value)
    di = ParseMDY(txtDateIssued.Value)

    If IsEmpty(bd) Then
        MsgBox "Birthdate is invalid.", vbExclamation
        Exit Function
    End If

    If IsEmpty(di) Then
        MsgBox "Date Issued is invalid.", vbExclamation
        Exit Function
    End If

    ValidateInputs = True

End Function
Private Sub cmdSave_Click()

    Dim tbl As ListObject
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Set tbl = GetRegistryTable()

    If Not ValidateInputs() Then Exit Sub

    If IDExists(txtID.Value) Then
        txtID.Value = NextID()
        Exit Sub
    End If

    '==========================================================
    ' STORE ORIGINAL COLUMN WIDTHS
    '==========================================================
    Dim colWidths() As Double
    ReDim colWidths(1 To tbl.ListColumns.Count)

    Dim i As Long

    For i = 1 To tbl.ListColumns.Count
        colWidths(i) = tbl.ListColumns(i).Range.ColumnWidth
    Next i

    '==========================================================
    ' UNPROTECT SHEET
    '==========================================================
    ws.Unprotect "erpatsystem"

    '==========================================================
    ' ADD NEW RECORD
    '==========================================================
    Dim newRow As ListRow

    Set newRow = tbl.ListRows.Add

    With newRow.Range

        .Cells(1, 1).Value = txtID.Value
        .Cells(1, 2).Value = txtSurname.Value
        .Cells(1, 3).Value = txtFirstname.Value
        .Cells(1, 4).Value = txtMiddlename.Value

        .Cells(1, 5).Value = ParseMDY(txtBirthdate.Value)

        .Cells(1, 7).Value = cboCivilStatus.Value
        .Cells(1, 8).Value = cboBarangay.Value

        .Cells(1, 9).Value = NormalizePhone(txtPContact.Value)
        .Cells(1, 10).Value = txtEmergency.Value
        .Cells(1, 11).Value = NormalizePhone(txtContact.Value)

        .Cells(1, 12).Value = txtOccupation.Value
        .Cells(1, 13).Value = ParseMDY(txtDateIssued.Value)

    End With

    '==========================================================
    ' RESTORE COLUMN WIDTHS
    '==========================================================
    For i = 1 To tbl.ListColumns.Count

        If colWidths(i) > 0 Then
            tbl.ListColumns(i).Range.ColumnWidth = colWidths(i)
        End If

    Next i

    '==========================================================
    ' REPROTECT SHEET
    '==========================================================
    ws.Protect Password:="erpatsystem", _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True, _
               AllowUsingPivotTables:=True

    ws.EnableSelection = xlNoRestrictions

    MsgBox "Added successfully!", vbInformation
    
    WriteSystemLog _
    "ADDED", _
    txtID.Value, _
    txtSurname.Value & ", " & txtFirstname.Value

    '==========================================================
    ' SAFE DASHBOARD REFRESH
    '==========================================================
    RefreshSystemAnalytics

    RefreshList txtSearch.Value
    StartNewRecord

End Sub


Private Sub cmdUpdate_Click()

    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim oldID As String
    Dim newID As String
    Dim f As Range

    Set tbl = GetRegistryTable()
    Set ws = tbl.Parent

    If mEditingRowIndex <= 0 Then Exit Sub
    If Not ValidateInputs() Then Exit Sub

    oldID = CStr(tbl.DataBodyRange.Cells(mEditingRowIndex, 1).Value)
    newID = Trim(txtID.Value)

    If newID <> oldID Then

        Set f = tbl.ListColumns(1).DataBodyRange.Find(newID, , xlValues, xlWhole)

        If Not f Is Nothing Then
            MsgBox "ID already exists!", vbExclamation
            txtID.Value = oldID
            Exit Sub
        End If

    End If

    '==========================================================
    ' STORE COLUMN WIDTHS
    '==========================================================
    Dim colWidths() As Double
    ReDim colWidths(1 To tbl.ListColumns.Count)

    Dim i As Long

    For i = 1 To tbl.ListColumns.Count
        colWidths(i) = tbl.ListColumns(i).Range.ColumnWidth
    Next i

    '==========================================================
    ' UNPROTECT
    '==========================================================
    ws.Unprotect Password:="erpatsystem"

    '==========================================================
    ' UPDATE RECORD
    '==========================================================
    With tbl.DataBodyRange

        .Cells(mEditingRowIndex, 1).Value = newID

        .Cells(mEditingRowIndex, 2).Value = txtSurname.Value
        .Cells(mEditingRowIndex, 3).Value = txtFirstname.Value
        .Cells(mEditingRowIndex, 4).Value = txtMiddlename.Value

        .Cells(mEditingRowIndex, 5).Value = ParseMDY(txtBirthdate.Value)

        .Cells(mEditingRowIndex, 7).Value = cboCivilStatus.Value
        .Cells(mEditingRowIndex, 8).Value = cboBarangay.Value

        .Cells(mEditingRowIndex, 9).Value = NormalizePhone(txtPContact.Value)
        .Cells(mEditingRowIndex, 10).Value = txtEmergency.Value
        .Cells(mEditingRowIndex, 11).Value = NormalizePhone(txtContact.Value)

        .Cells(mEditingRowIndex, 12).Value = txtOccupation.Value
        .Cells(mEditingRowIndex, 13).Value = ParseMDY(txtDateIssued.Value)

    End With

    '==========================================================
    ' RESTORE COLUMN WIDTHS
    '==========================================================
    For i = 1 To tbl.ListColumns.Count

        If colWidths(i) > 0 Then
            tbl.ListColumns(i).Range.ColumnWidth = colWidths(i)
        End If

    Next i

    '==========================================================
    ' REPROTECT
    '==========================================================
    ws.Protect Password:="erpatsystem", _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True, _
               UserInterfaceOnly:=True, _
               AllowFiltering:=True, _
               AllowSorting:=True, _
               AllowUsingPivotTables:=True

    ws.EnableSelection = xlNoRestrictions

    MsgBox "Updated successfully!", vbInformation
    
    WriteSystemLog _
    "ADDED", _
    txtID.Value, _
    txtSurname.Value & ", " & txtFirstname.Value

    '==========================================================
    ' SAFE DASHBOARD REFRESH
    '==========================================================
    RefreshSystemAnalytics

    RefreshList txtSearch.Value

End Sub

Private Sub cmdDelete_Click()

    Dim tbl As ListObject
    Dim ws As Worksheet

    Set tbl = GetRegistryTable()
    Set ws = tbl.Parent

    If mEditingRowIndex <= 0 Then Exit Sub

    If MsgBox("Delete this record?", vbYesNo + vbQuestion) = vbYes Then

        '======================================================
        ' STORE COLUMN WIDTHS
        '======================================================
        Dim colWidths() As Double
        ReDim colWidths(1 To tbl.ListColumns.Count)

        Dim i As Long

        For i = 1 To tbl.ListColumns.Count
            colWidths(i) = tbl.ListColumns(i).Range.ColumnWidth
        Next i

        '======================================================
        ' UNPROTECT SHEET
        '======================================================
        ws.Unprotect Password:="erpatsystem"

        '======================================================
        ' DELETE ROW
        '======================================================
        Dim deletedName As String
Dim deletedID As String

deletedID = tbl.DataBodyRange.Cells(mEditingRowIndex, 1).Value

deletedName = _
    tbl.DataBodyRange.Cells(mEditingRowIndex, 2).Value & ", " & _
    tbl.DataBodyRange.Cells(mEditingRowIndex, 3).Value
    
        tbl.ListRows(mEditingRowIndex).Delete

        '======================================================
        ' RESTORE WIDTHS
        '======================================================
        For i = 1 To tbl.ListColumns.Count

            If colWidths(i) > 0 Then
                tbl.ListColumns(i).Range.ColumnWidth = colWidths(i)
            End If

        Next i

        '======================================================
        ' REPROTECT SHEET
        '======================================================
        ws.Protect Password:="erpatsystem", _
                   DrawingObjects:=True, _
                   Contents:=True, _
                   Scenarios:=True, _
                   UserInterfaceOnly:=True, _
                   AllowFiltering:=True, _
                   AllowSorting:=True, _
                   AllowUsingPivotTables:=True

        ws.EnableSelection = xlNoRestrictions

        MsgBox "Deleted.", vbInformation
        
        WriteSystemLog _
    "DELETED", _
    deletedID, _
    deletedName

        '======================================================
        ' SAFE DASHBOARD REFRESH
        '======================================================
        RefreshSystemAnalytics

        RefreshList txtSearch.Value
        StartNewRecord

    End If

End Sub





Private Function FormatCellDate(ByVal v As Variant) As String

    If IsDate(v) Then
        FormatCellDate = Format(v, "mm/dd/yyyy")
    Else
        FormatCellDate = ""
    End If

End Function

Private Function ParseMDY(ByVal s As String) As Variant

    Dim p() As String

    s = Trim(s)

    If s = "" Then
        ParseMDY = Empty
        Exit Function
    End If

    p = Split(s, "/")

    If UBound(p) <> 2 Then
        ParseMDY = Empty
        Exit Function
    End If

    On Error GoTo bad

    ParseMDY = DateSerial(CInt(p(2)), CInt(p(0)), CInt(p(1)))
    Exit Function

bad:
    ParseMDY = Empty

End Function


