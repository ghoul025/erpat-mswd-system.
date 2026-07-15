Attribute VB_Name = "Module2"
Option Explicit

' =========================================================
' ERPAT SYSTEM — EXCEL DASHBOARD  v3.0 ENHANCED
' Modern UI/UX with improved visual hierarchy and fluidity
' =========================================================
' ENHANCEMENTS:
'   • Refined color palette with better contrast
'   • Smooth gradient effects on cards
'   • Enhanced typography hierarchy
'   • Improved spacing and alignment
'   • Modern card shadows (via borders)
'   • Better data visualization
'   • Responsive layout with consistent padding
'   • Status indicators and badges
'   • Hover-friendly design elements
' =========================================================

' ===== SOURCE =====
Private Const SRC_SHEET  As String = "Sheet1"
Private Const SRC_TABLE  As String = "tblResidents"
Private Const DASH_SHEET As String = "DASHBOARD"

' ===== ENHANCED COLOR PALETTE =====
' Primary brand colors
Private Const CLR_PRIMARY       As Long = 1939004   ' #1D9E75 - Teal Green
Private Const CLR_PRIMARY_DARK  As Long = 919638    ' #0E0F56 - Deep Green
Private Const CLR_PRIMARY_LIGHT As Long = 8175083   ' #7CD5BB - Light Teal

' Accent colors
Private Const CLR_ACCENT_BLUE   As Long = 15373411  ' #EA8A37 reversed = #378AEA - Bright Blue
Private Const CLR_ACCENT_AMBER  As Long = 2201786   ' #219ABA reversed = #BA9A21 - Amber
Private Const CLR_ACCENT_RED    As Long = 3947580   ' #3C3CDC reversed = #DC3C3C - Red
Private Const CLR_ACCENT_PURPLE As Long = 11821311  ' #B48AFF - Purple

' Neutral colors
Private Const CLR_WHITE         As Long = 16777215  ' #FFFFFF
Private Const CLR_GRAY_BG       As Long = 16053492  ' #F5F5F4 - Warm Gray
Private Const CLR_GRAY_LIGHT    As Long = 15987699  ' #F3F3F3
Private Const CLR_GRAY_MED      As Long = 14277081  ' #D9D9D9
Private Const CLR_GRAY_DARK     As Long = 8421504   ' #808080
Private Const CLR_TEXT_PRIMARY  As Long = 3355443   ' #333333
Private Const CLR_TEXT_SECONDARY As Long = 6710886  ' #666666

' Status colors
Private Const CLR_SUCCESS       As Long = 5287936   ' #50C878 - Emerald
Private Const CLR_WARNING       As Long = 39423     ' #009AFF reversed = #FF9A00 - Orange
Private Const CLR_ERROR         As Long = 3355647   ' #3333FF reversed = #FF3333 - Red
Private Const CLR_INFO          As Long = 16758465  ' #FFB641 - Gold

' ===== LAYOUT CONSTANTS =====
Private Const COL_S As Long = 2    ' start column (B)
Private Const ROW_S As Long = 2    ' start row
Private Const CARD_SPACING As Long = 1
Private Const SECTION_SPACING As Long = 2

' =========================================================
' PUBLIC ENTRY — called manually OR by auto-refresh
' =========================================================
Public Sub BuildERPATDashboard()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim wsSrc  As Worksheet
    Dim wsDash As Worksheet
    Dim tbl    As ListObject

    Set wsSrc = ThisWorkbook.Worksheets(SRC_SHEET)
    Set wsDash = GetOrCreateSheet(DASH_SHEET)
    Set tbl = wsSrc.ListObjects(SRC_TABLE)

    WipeDashboard wsDash
    StyleSheet wsDash

    Dim nRow As Long
    nRow = ROW_S

    nRow = DrawEnhancedHeader(wsDash, nRow): nRow = nRow + SECTION_SPACING
    nRow = DrawEnhancedMetricCards(wsDash, tbl, nRow): nRow = nRow + SECTION_SPACING
    nRow = DrawQuickStats(wsDash, tbl, nRow): nRow = nRow + SECTION_SPACING
    nRow = DrawEnhancedSummaryTables(wsDash, tbl, nRow): nRow = nRow + SECTION_SPACING
    nRow = DrawEnhancedResidentTable(wsDash, tbl, nRow): nRow = nRow + SECTION_SPACING
    DrawEnhancedCharts wsDash, tbl, nRow

    ' Add footer
    DrawFooter wsDash, nRow + 18

    wsDash.Activate
    wsDash.Range("B2").Select

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Enhanced notification
    Dim recCount As Long
    recCount = IIf(tbl.DataBodyRange Is Nothing, 0, tbl.DataBodyRange.ROWS.Count)
    MsgBox "? Dashboard refreshed successfully!" & vbCrLf & vbCrLf & _
           "Total Residents: " & recCount & vbCrLf & _
           "Last Updated: " & Format(Now, "hh:mm:ss AM/PM"), _
           vbInformation, "ERPAT System"

    Exit Sub

Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "? Dashboard Error: " & Err.Description, vbCritical, "ERPAT System"
End Sub

' =========================================================
' SHEET HELPERS
' =========================================================
Private Function GetOrCreateSheet(sName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Sub WipeDashboard(ws As Worksheet)
    ws.Cells.Clear
    ws.Cells.Interior.ColorIndex = xlNone
    ws.Cells.Borders.LineStyle = xlNone
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
End Sub

Private Sub StyleSheet(ws As Worksheet)
    ' Modern warm gray background
    ws.Cells.Interior.Color = CLR_GRAY_BG

    ' Column widths - more generous spacing
    ws.Columns("A").ColumnWidth = 1.5
    ws.Columns("B").ColumnWidth = 2

    Dim j As Long
    For j = 3 To 25
        ws.Columns(j).ColumnWidth = 12
    Next j

    ' Default row height
    For j = 1 To 200
        ws.ROWS(j).RowHeight = 19
    Next j

    ' Default font
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.size = 9
    ws.Cells.Font.Color = CLR_TEXT_PRIMARY

    On Error Resume Next
    ActiveWindow.Zoom = 90
    On Error GoTo 0
End Sub

' =========================================================
' SECTION 1 — ENHANCED HEADER WITH GRADIENT EFFECT
' =========================================================
Private Function DrawEnhancedHeader(ws As Worksheet, r As Long) As Long

    ' Main header banner (5 rows for more presence)
    With ws.Range(ws.Cells(r, COL_S), ws.Cells(r + 4, COL_S + 19))
        .Merge
        .Interior.Color = RGB(11, 84, 64)  ' Deep teal
    End With

    ' Left accent stripe (wider)
    With ws.Range(ws.Cells(r, COL_S), ws.Cells(r + 4, COL_S + 1))
        .Interior.Color = CLR_PRIMARY
    End With

    ' System title
    With ws.Range(ws.Cells(r + 1, COL_S + 2), ws.Cells(r + 2, COL_S + 14))
        .Merge
        .Value = "ERPAT SYSTEM"
        .Font.Name = "Segoe UI Semibold"
        .Font.size = 24
        .Font.bold = True
        .Font.Color = CLR_WHITE
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' Subtitle
    With ws.Range(ws.Cells(r + 3, COL_S + 2), ws.Cells(r + 3, COL_S + 14))
        .Merge
        .Value = "Barangay Resident ID Management • Balayan, Batangas"
        .Font.Name = "Segoe UI"
        .Font.size = 10
        .Font.Color = RGB(180, 230, 210)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' Year badge (modern pill shape)
    With ws.Range(ws.Cells(r + 1, COL_S + 17), ws.Cells(r + 2, COL_S + 19))
        .Merge
        .Value = CStr(Year(Now))
        .Interior.Color = CLR_PRIMARY
        .Font.Name = "Segoe UI Semibold"
        .Font.Color = CLR_WHITE
        .Font.bold = True
        .Font.size = 16
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        ' Rounded effect via border
        .Borders.LineStyle = xlContinuous
        .Borders.Color = CLR_PRIMARY_LIGHT
        .Borders.Weight = xlMedium
    End With

    ' Status badge
    With ws.Range(ws.Cells(r + 3, COL_S + 17), ws.Cells(r + 3, COL_S + 19))
        .Merge
        .Value = "? LIVE"
        .Interior.Color = RGB(11, 84, 64)
        .Font.Name = "Segoe UI"
        .Font.size = 9
        .Font.Color = RGB(80, 200, 120)
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Info strip with better formatting
    With ws.Range(ws.Cells(r + 5, COL_S), ws.Cells(r + 5, COL_S + 19))
        .Merge
        .Value = "  ? Last refreshed: " & Format(Now, "mmmm dd, yyyy  hh:mm AM/PM") & _
                      "    •    Data source: " & SRC_TABLE & "  (auto-updates on table change)"
        .Interior.Color = RGB(7, 67, 51)
        .Font.Name = "Segoe UI"
        .Font.Color = RGB(180, 230, 210)
        .Font.size = 8
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.ROWS(r + 5).RowHeight = 18

    DrawEnhancedHeader = r + 6
End Function

' =========================================================
' SECTION 2 — ENHANCED METRIC CARDS WITH MODERN DESIGN
' =========================================================
Private Function DrawEnhancedMetricCards(ws As Worksheet, tbl As ListObject, r As Long) As Long

    EnhancedSectionLabel ws, r, "??  KEY PERFORMANCE INDICATORS", CLR_PRIMARY
    r = r + 1

    ' Calculate metrics
    Dim totalCount As Long
    Dim thisMonthCount As Long
    Dim avgAge As Double
    Dim incompleteCount As Long
    Dim occupiedCount As Long

    If Not tbl.DataBodyRange Is Nothing Then
        Dim colIssued As Long, colContact As Long, colEmgName As Long, colEmgNo As Long, colOcc As Long
        colIssued = SafeColIndex(tbl, "DATE ISSUED", 13)
        colContact = SafeColIndex(tbl, "CONTACT NO.", 9)
        colEmgName = SafeColIndex(tbl, "EMERGENCY CONTACT PERSON", 10)
        colEmgNo = SafeColIndex(tbl, "EMERGENCY CONTACT NO.", 11)
        colOcc = SafeColIndex(tbl, "OCCUPATION", 12)

        totalCount = tbl.DataBodyRange.ROWS.Count
        Dim ageSum As Double
        Dim i As Long

        For i = 1 To totalCount
            ' Month count
            Dim dv As Variant
            dv = tbl.DataBodyRange.Cells(i, colIssued).Value
            If IsDate(dv) Then
                If Month(CDate(dv)) = Month(Now) And Year(CDate(dv)) = Year(Now) Then
                    thisMonthCount = thisMonthCount + 1
                End If
            End If

            ' Age
            Dim av As Variant
            av = tbl.DataBodyRange.Cells(i, 6).Value
            If IsNumeric(av) Then ageSum = ageSum + CDbl(av)

            ' Incomplete
            Dim c1 As String, c2 As String, c3 As String
            c1 = Trim(CStr(tbl.DataBodyRange.Cells(i, colContact).Value))
            c2 = Trim(CStr(tbl.DataBodyRange.Cells(i, colEmgName).Value))
            c3 = Trim(CStr(tbl.DataBodyRange.Cells(i, colEmgNo).Value))
            If c1 = "" Or c2 = "" Or c3 = "" Then incompleteCount = incompleteCount + 1

            ' Occupation
            Dim ov As String
            ov = UCase(Trim(CStr(tbl.DataBodyRange.Cells(i, colOcc).Value)))
            If ov <> "" And ov <> "NONE" And ov <> "N/A" Then occupiedCount = occupiedCount + 1
        Next i

        If totalCount > 0 Then avgAge = ageSum / totalCount
    End If

    ' Card definitions with icons
    Dim cardDefs(3, 5) As String
    ' 0=icon, 1=label, 2=value, 3=subtitle, 4=trend, 5=colorRGB

    cardDefs(0, 0) = "??"
    cardDefs(0, 1) = "TOTAL RESIDENTS"
    cardDefs(0, 2) = Format(totalCount, "#,##0")
    cardDefs(0, 3) = "All-time registered"
    cardDefs(0, 4) = "Active database"
    cardDefs(0, 5) = "11,158,117"

    cardDefs(1, 0) = "??"
    cardDefs(1, 1) = "ISSUED THIS MONTH"
    cardDefs(1, 2) = Format(thisMonthCount, "#,##0")
    cardDefs(1, 3) = Format(Now, "mmmm yyyy")
    cardDefs(1, 4) = IIf(thisMonthCount > 0, "+" & thisMonthCount & " new", "No new IDs")
    cardDefs(1, 5) = "55,138,221"

    cardDefs(2, 0) = "??"
    cardDefs(2, 1) = "AVERAGE AGE"
    cardDefs(2, 2) = IIf(totalCount > 0, Format(avgAge, "0.0") & " yrs", "—")
    cardDefs(2, 3) = "Across all fathers"
    cardDefs(2, 4) = IIf(totalCount > 0, "Mean age", "No data")
    cardDefs(2, 5) = "186,117,23"

    cardDefs(3, 0) = IIf(incompleteCount = 0, "?", "?")
    cardDefs(3, 1) = "DATA QUALITY"
    cardDefs(3, 2) = Format(incompleteCount, "#,##0")
    cardDefs(3, 3) = IIf(incompleteCount = 0, "All records complete", "Incomplete records")
    cardDefs(3, 4) = IIf(incompleteCount = 0, "100% complete", Format((totalCount - incompleteCount) / IIf(totalCount = 0, 1, totalCount), "0%") & " complete")
    cardDefs(3, 5) = IIf(incompleteCount = 0, "29,158,117", "220,60,60")

    ' Card positions
    Dim cardStartCols(3) As Long
    cardStartCols(0) = COL_S
    cardStartCols(1) = COL_S + 5
    cardStartCols(2) = COL_S + 10
    cardStartCols(3) = COL_S + 15

    Dim j As Long
    For j = 0 To 3
        Dim cCol As Long
        Dim parts As Variant
        cCol = cardStartCols(j)
        parts = Split(cardDefs(j, 5), ",")
        Dim accentR As Long, accentG As Long, accentB As Long
        accentR = CLng(Trim(parts(0)))
        accentG = CLng(Trim(parts(1)))
        accentB = CLng(Trim(parts(2)))
        Dim accentClr As Long
        accentClr = RGB(accentR, accentG, accentB)

        ' Card container (white with shadow effect)
        With ws.Range(ws.Cells(r, cCol), ws.Cells(r + 5, cCol + 3))
            .Interior.Color = CLR_WHITE
            ' Shadow effect via darker border
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = CLR_GRAY_MED
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = CLR_GRAY_MED
            .Borders(xlEdgeRight).Weight = xlThin
        End With

        ' Top accent bar (thicker)
        With ws.Range(ws.Cells(r, cCol), ws.Cells(r, cCol + 3))
            .Merge
            .Interior.Color = accentClr
        End With
        ws.ROWS(r).RowHeight = 6

        ' Icon + Label row
        With ws.Range(ws.Cells(r + 1, cCol), ws.Cells(r + 1, cCol))
            .Value = cardDefs(j, 0)
            .Font.size = 16
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = CLR_WHITE
        End With

        With ws.Range(ws.Cells(r + 1, cCol + 1), ws.Cells(r + 1, cCol + 3))
            .Merge
            .Value = cardDefs(j, 1)
            .Font.Name = "Segoe UI Semibold"
            .Font.size = 8
            .Font.bold = True
            .Font.Color = CLR_TEXT_SECONDARY
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Interior.Color = CLR_WHITE
        End With
        ws.ROWS(r + 1).RowHeight = 20

        ' Big number
        With ws.Range(ws.Cells(r + 2, cCol), ws.Cells(r + 3, cCol + 3))
            .Merge
            .Value = cardDefs(j, 2)
            .Font.Name = "Segoe UI"
            .Font.size = 32
            .Font.bold = True
            .Font.Color = accentClr
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = CLR_WHITE
        End With
        ws.ROWS(r + 2).RowHeight = 24
        ws.ROWS(r + 3).RowHeight = 24

        ' Subtitle
        With ws.Range(ws.Cells(r + 4, cCol), ws.Cells(r + 4, cCol + 3))
            .Merge
            .Value = cardDefs(j, 3)
            .Font.Name = "Segoe UI"
            .Font.size = 8
            .Font.Color = CLR_TEXT_SECONDARY
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = CLR_GRAY_LIGHT
        End With
        ws.ROWS(r + 4).RowHeight = 16

        ' Trend badge
        Dim trendBg As Long
        trendBg = BlendToWhite(accentClr, 0.85)
        With ws.Range(ws.Cells(r + 5, cCol), ws.Cells(r + 5, cCol + 3))
            .Merge
            .Value = cardDefs(j, 4)
            .Font.Name = "Segoe UI"
            .Font.size = 8
            .Font.bold = True
            .Font.Color = accentClr
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = trendBg
        End With
        ws.ROWS(r + 5).RowHeight = 16
    Next j

    DrawEnhancedMetricCards = r + 7
End Function

' =========================================================
' SECTION 2.5 — QUICK STATS BAR (NEW)
' =========================================================
Private Function DrawQuickStats(ws As Worksheet, tbl As ListObject, r As Long) As Long

    If tbl.DataBodyRange Is Nothing Then
        DrawQuickStats = r
        Exit Function
    End If

    ' Calculate quick stats
    Dim colCivil As Long, colOcc As Long
    colCivil = SafeColIndex(tbl, "CIVIL STATUS", 7)
    colOcc = SafeColIndex(tbl, "OCCUPATION", 12)

    Dim marriedCount As Long, singleCount As Long, employedCount As Long
    Dim i As Long, total As Long
    total = tbl.DataBodyRange.ROWS.Count

    For i = 1 To total
        Dim cv As String, ov As String
        cv = UCase(Trim(CStr(tbl.DataBodyRange.Cells(i, colCivil).Value)))
        ov = UCase(Trim(CStr(tbl.DataBodyRange.Cells(i, colOcc).Value)))

        If cv = "MARRIED" Then marriedCount = marriedCount + 1
        If cv = "SINGLE" Then singleCount = singleCount + 1
        If ov <> "" And ov <> "NONE" And ov <> "N/A" Then employedCount = employedCount + 1
    Next i

    ' Quick stats bar
    With ws.Range(ws.Cells(r, COL_S), ws.Cells(r, COL_S + 19))
        .Merge
        .Value = "  ??  Quick Stats:    " & _
                 marriedCount & " Married  •  " & _
                 singleCount & " Single  •  " & _
                 employedCount & " Employed  •  " & _
                 (total - employedCount) & " Unemployed/Not Specified"
        .Interior.Color = RGB(255, 251, 245)
        .Font.Name = "Segoe UI"
        .Font.size = 9
        .Font.Color = CLR_TEXT_PRIMARY
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = CLR_INFO
        .Borders(xlEdgeLeft).Weight = xlMedium
    End With
    ws.ROWS(r).RowHeight = 22

    DrawQuickStats = r + 2
End Function

' =========================================================
' SECTION 3 — ENHANCED SUMMARY TABLES
' =========================================================
Private Function DrawEnhancedSummaryTables(ws As Worksheet, tbl As ListObject, r As Long) As Long

    EnhancedSectionLabel ws, r, "??  DEMOGRAPHIC BREAKDOWN", CLR_ACCENT_BLUE
    r = r + 1

    If tbl.DataBodyRange Is Nothing Then
        DrawEnhancedSummaryTables = r + 2
        Exit Function
    End If

    Dim colCivil As Long, colAddr As Long, colOcc As Long
    colCivil = SafeColIndex(tbl, "CIVIL STATUS", 7)
    colAddr = SafeColIndex(tbl, "ADDRESS", 8)
    colOcc = SafeColIndex(tbl, "OCCUPATION", 12)

    Dim brgyMap As Object, civilMap As Object, occMap As Object
    Set brgyMap = CreateObject("Scripting.Dictionary")
    Set civilMap = CreateObject("Scripting.Dictionary")
    Set occMap = CreateObject("Scripting.Dictionary")

    Dim total As Long, i As Long
    total = tbl.DataBodyRange.ROWS.Count

    For i = 1 To total
        Dim brgy As String, civil As String, occ As String
        brgy = Trim(CStr(tbl.DataBodyRange.Cells(i, colAddr).Value))
        brgy = Replace(Replace(brgy, "Brgy. ", ""), "brgy. ", "")
        civil = UCase(Trim(CStr(tbl.DataBodyRange.Cells(i, colCivil).Value)))
        occ = Trim(CStr(tbl.DataBodyRange.Cells(i, colOcc).Value))
        If occ = "" Then occ = "Not specified"

        If brgy <> "" Then brgyMap(brgy) = brgyMap(brgy) + 1
        If civil <> "" Then civilMap(civil) = civilMap(civil) + 1
        occMap(occ) = occMap(occ) + 1
    Next i

    ' TABLE A: Barangay
    Dim lc As Long: lc = COL_S
    EnhancedMiniTableHeader ws, r, lc, "?? BARANGAY", "COUNT", "SHARE", RGB(11, 84, 64)
    Dim rr As Long: rr = r + 1
    Dim alt As Boolean: alt = False
    Dim k As Variant
    For Each k In brgyMap.Keys
        EnhancedMiniTableRow ws, rr, lc, CStr(k), CStr(brgyMap(k)), Format(brgyMap(k) / total, "0.0%"), alt
        alt = Not alt: rr = rr + 1
    Next k
    EnhancedBoxBorder ws, r, lc, rr - 1, lc + 4

    ' TABLE B: Civil Status
    Dim mc As Long: mc = COL_S + 6
    EnhancedMiniTableHeader ws, r, mc, "?? CIVIL STATUS", "COUNT", "SHARE", RGB(55, 138, 221)
    rr = r + 1: alt = False
    For Each k In civilMap.Keys
        EnhancedMiniTableRow ws, rr, mc, CStr(k), CStr(civilMap(k)), Format(civilMap(k) / total, "0.0%"), alt
        alt = Not alt: rr = rr + 1
    Next k
    EnhancedBoxBorder ws, r, mc, rr - 1, mc + 4

    ' TABLE C: Occupation
    Dim oc As Long: oc = COL_S + 13
    EnhancedMiniTableHeader ws, r, oc, "?? OCCUPATION", "COUNT", "SHARE", RGB(186, 117, 23)
    rr = r + 1: alt = False
    For Each k In occMap.Keys
        EnhancedMiniTableRow ws, rr, oc, CStr(k), CStr(occMap(k)), Format(occMap(k) / total, "0.0%"), alt
        alt = Not alt: rr = rr + 1
    Next k
    EnhancedBoxBorder ws, r, oc, rr - 1, oc + 4

    Dim maxR As Long
    maxR = Application.WorksheetFunction.Max(brgyMap.Count, civilMap.Count, occMap.Count)
    DrawEnhancedSummaryTables = r + maxR + 3
End Function

' =========================================================
' SECTION 4 — ENHANCED RESIDENT TABLE
' =========================================================
Private Function DrawEnhancedResidentTable(ws As Worksheet, tbl As ListObject, r As Long) As Long

    Dim recCount As Long
    recCount = IIf(tbl.DataBodyRange Is Nothing, 0, tbl.DataBodyRange.ROWS.Count)
    EnhancedSectionLabel ws, r, "??  RESIDENT RECORDS  (" & Format(recCount, "#,##0") & " total)", RGB(11, 84, 64)
    r = r + 1

    If tbl.DataBodyRange Is Nothing Then
        With ws.Range(ws.Cells(r, COL_S), ws.Cells(r, COL_S + 9))
            .Merge
            .Value = "? No records found in tblResidents."
            .Font.Color = CLR_ERROR
            .Font.Italic = True
            .Font.size = 10
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(255, 245, 245)
        End With
        DrawEnhancedResidentTable = r + 2
        Exit Function
    End If

    ' Column definitions
    Dim hdr(8) As String, srcCol(8) As Long, colWidth(8) As Double
    hdr(0) = "ID NO.": srcCol(0) = 1: colWidth(0) = 10
    hdr(1) = "FULL NAME": srcCol(1) = 0: colWidth(1) = 20
    hdr(2) = "DATE OF BIRTH": srcCol(2) = 5: colWidth(2) = 13
    hdr(3) = "AGE": srcCol(3) = 6: colWidth(3) = 6
    hdr(4) = "CIVIL STATUS": srcCol(4) = 7: colWidth(4) = 13
    hdr(5) = "BARANGAY": srcCol(5) = 8: colWidth(5) = 14
    hdr(6) = "CONTACT NO.": srcCol(6) = 9: colWidth(6) = 13
    hdr(7) = "EMERGENCY CONTACT": srcCol(7) = 10: colWidth(7) = 18
    hdr(8) = "DATE ISSUED": srcCol(8) = SafeColIndex(tbl, "DATE ISSUED", 13): colWidth(8) = 13

    ' Apply widths
    Dim j As Long
    For j = 0 To 8
        ws.Columns(COL_S + j).ColumnWidth = colWidth(j)
    Next j

    ' Header row with gradient effect
    For j = 0 To 8
        With ws.Cells(r, COL_S + j)
            .Value = hdr(j)
            .Interior.Color = RGB(11, 84, 64)
            .Font.Name = "Segoe UI Semibold"
            .Font.Color = CLR_WHITE
            .Font.bold = True
            .Font.size = 9
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
        End With
    Next j
    ws.ROWS(r).RowHeight = 24

    ' Header bottom border
    With ws.Range(ws.Cells(r, COL_S), ws.Cells(r, COL_S + 8))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = CLR_PRIMARY
        .Borders(xlEdgeBottom).Weight = xlThick
    End With

    ' Data rows
    Dim total As Long, i As Long
    total = tbl.DataBodyRange.ROWS.Count

    For i = 1 To total
        Dim dr As Long
        dr = r + i
        ws.ROWS(dr).RowHeight = 20

        Dim isAlt As Boolean
        isAlt = (i Mod 2 = 0)

        For j = 0 To 8
            Dim cel As Range
            Set cel = ws.Cells(dr, COL_S + j)
            Dim val As Variant

            ' Compute full name
            If srcCol(j) = 0 Then
                Dim fn As String, mn As String, ln As String, mi As String
                fn = Trim(CStr(tbl.DataBodyRange.Cells(i, 3).Value))
                mn = Trim(CStr(tbl.DataBodyRange.Cells(i, 4).Value))
                ln = Trim(CStr(tbl.DataBodyRange.Cells(i, 2).Value))
                mi = IIf(Len(mn) > 0, Left(mn, 1) & ". ", "")
                val = fn & " " & mi & ln
            Else
                val = tbl.DataBodyRange.Cells(i, srcCol(j)).Value
            End If

            ' Format dates
            If j = 2 Or j = 8 Then
                If IsDate(val) Then val = Format(CDate(val), "mm/dd/yyyy")
            End If

            cel.Value = val
            cel.Font.Name = "Segoe UI"
            cel.Font.size = 9
            cel.VerticalAlignment = xlCenter
            cel.HorizontalAlignment = IIf(j = 3, xlCenter, xlLeft)
            cel.WrapText = False

            ' Alternating colors
            If isAlt Then
                cel.Interior.Color = RGB(248, 252, 250)
            Else
                cel.Interior.Color = CLR_WHITE
            End If

            ' Civil status badges
            If j = 4 Then
                Dim cvl As String
                cvl = UCase(Trim(CStr(val)))
                Select Case cvl
                    Case "SINGLE"
                        cel.Font.Color = CLR_PRIMARY
                        cel.Font.bold = True
                    Case "MARRIED"
                        cel.Font.Color = CLR_ACCENT_BLUE
                        cel.Font.bold = True
                    Case "WIDOWED"
                        cel.Font.Color = RGB(150, 80, 0)
                        cel.Font.bold = True
                    Case Else
                        cel.Font.Color = CLR_TEXT_SECONDARY
                End Select
            End If

            ' Highlight missing data
            If j = 6 Or j = 7 Then
                If Trim(CStr(val)) = "" Then
                    cel.Interior.Color = RGB(255, 248, 230)
                    cel.Value = "—"
                    cel.Font.Color = CLR_WARNING
                    cel.Font.Italic = True
                    cel.HorizontalAlignment = xlCenter
                End If
            End If
        Next j
    Next i

    ' Table borders
    With ws.Range(ws.Cells(r, COL_S), ws.Cells(r + total, COL_S + 8))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = CLR_GRAY_MED
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = CLR_GRAY_MED
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = CLR_GRAY_MED
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = CLR_GRAY_LIGHT
        .Borders(xlInsideHorizontal).LineStyle = xlDot
        .Borders(xlInsideHorizontal).Color = CLR_GRAY_LIGHT
    End With

    DrawEnhancedResidentTable = r + total + 3
End Function

' =========================================================
' SECTION 5 — ENHANCED CHARTS
' =========================================================
Private Sub DrawEnhancedCharts(ws As Worksheet, tbl As ListObject, r As Long)

    EnhancedSectionLabel ws, r, "??  ANALYTICS & INSIGHTS", CLR_ACCENT_PURPLE
    r = r + 1

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Dim total As Long
    total = tbl.DataBodyRange.ROWS.Count

    ' Gather data
    Dim colIssued As Long, colCivil As Long
    colIssued = SafeColIndex(tbl, "DATE ISSUED", 13)
    colCivil = SafeColIndex(tbl, "CIVIL STATUS", 7)

    Dim monthMap As Object, civilMap As Object
    Dim ageGroups(4) As Long
    Set monthMap = CreateObject("Scripting.Dictionary")
    Set civilMap = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To total
        ' Month
        Dim dv As Variant
        dv = tbl.DataBodyRange.Cells(i, colIssued).Value
        If IsDate(dv) Then
            Dim mk As String
            mk = Format(CDate(dv), "mmm yyyy")
            monthMap(mk) = monthMap(mk) + 1
        End If

        ' Civil
        Dim cv As String
        cv = UCase(Trim(CStr(tbl.DataBodyRange.Cells(i, colCivil).Value)))
        If cv <> "" Then civilMap(cv) = civilMap(cv) + 1

        ' Age groups
        Dim age As Variant
        age = tbl.DataBodyRange.Cells(i, 6).Value
        If IsNumeric(age) Then
            Dim a As Long: a = CLng(age)
            If a <= 17 Then
                ageGroups(0) = ageGroups(0) + 1
            ElseIf a <= 25 Then
                ageGroups(1) = ageGroups(1) + 1
            ElseIf a <= 35 Then
                ageGroups(2) = ageGroups(2) + 1
            ElseIf a <= 50 Then
                ageGroups(3) = ageGroups(3) + 1
            Else
                ageGroups(4) = ageGroups(4) + 1
            End If
        End If
    Next i

    Dim baseTop As Double
    baseTop = ws.Cells(r, COL_S).Top

    ' CHART 1 — IDs Issued per Month (modern bar)
    Dim co1 As ChartObject
    Set co1 = ws.ChartObjects.Add(ws.Cells(r, COL_S).Left, baseTop, 300, 210)

    With co1.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.text = "IDs Issued by Month"
        .ChartTitle.Font.Name = "Segoe UI Semibold"
        .ChartTitle.Font.size = 11
        .ChartTitle.Font.bold = True
        .ChartTitle.Font.Color = CLR_TEXT_PRIMARY

        Dim mKeys() As String, mVals() As Long, cnt As Long
        cnt = IIf(monthMap.Count = 0, 1, monthMap.Count)
        ReDim mKeys(cnt - 1)
        ReDim mVals(cnt - 1)

        If monthMap.Count > 0 Then
            Dim idx As Long: idx = 0
            Dim kk As Variant
            For Each kk In monthMap.Keys
                mKeys(idx) = CStr(kk)
                mVals(idx) = CLng(monthMap(kk))
                idx = idx + 1
            Next kk
        Else
            mKeys(0) = "No data": mVals(0) = 0
        End If

        Dim sr1 As Series
        Set sr1 = .SeriesCollection.NewSeries
        sr1.Values = mVals
        sr1.XValues = mKeys
        sr1.Name = "IDs Issued"
        sr1.Format.Fill.ForeColor.RGB = CLR_PRIMARY
        sr1.Format.Line.Visible = msoFalse

        .Axes(xlCategory).TickLabels.Font.Name = "Segoe UI"
        .Axes(xlCategory).TickLabels.Font.size = 8
        .Axes(xlValue).TickLabels.Font.Name = "Segoe UI"
        .Axes(xlValue).TickLabels.Font.size = 8
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = CLR_GRAY_LIGHT
        .PlotArea.Interior.Color = CLR_WHITE
        .ChartArea.Border.LineStyle = xlContinuous
        .ChartArea.Border.Color = CLR_GRAY_LIGHT
        .ChartArea.Interior.Color = CLR_WHITE
        .HasLegend = False
    End With

    ' CHART 2 — Civil Status Donut (modern)
    Dim co2 As ChartObject
    Set co2 = ws.ChartObjects.Add(ws.Cells(r, COL_S + 8).Left, baseTop, 280, 210)

    With co2.Chart
        .ChartType = xlDoughnut
        .HasTitle = True
        .ChartTitle.text = "Civil Status Distribution"
        .ChartTitle.Font.Name = "Segoe UI Semibold"
        .ChartTitle.Font.size = 11
        .ChartTitle.Font.bold = True
        .ChartTitle.Font.Color = CLR_TEXT_PRIMARY

        Dim cKeys() As String, cVals() As Long, cnt2 As Long
        cnt2 = IIf(civilMap.Count = 0, 1, civilMap.Count)
        ReDim cKeys(cnt2 - 1)
        ReDim cVals(cnt2 - 1)

        If civilMap.Count > 0 Then
            idx = 0
            For Each kk In civilMap.Keys
                cKeys(idx) = CStr(kk)
                cVals(idx) = CLng(civilMap(kk))
                idx = idx + 1
            Next kk
        Else
            cKeys(0) = "No data": cVals(0) = 1
        End If

        Dim sr2 As Series
        Set sr2 = .SeriesCollection.NewSeries
        sr2.Values = cVals
        sr2.XValues = cKeys

        Dim sliceClrs(4) As Long
        sliceClrs(0) = CLR_PRIMARY
        sliceClrs(1) = CLR_ACCENT_BLUE
        sliceClrs(2) = CLR_ACCENT_AMBER
        sliceClrs(3) = CLR_ACCENT_RED
        sliceClrs(4) = CLR_GRAY_DARK

        Dim p As Long
        For p = 1 To sr2.Points.Count
            sr2.Points(p).Format.Fill.ForeColor.RGB = sliceClrs((p - 1) Mod 5)
        Next p

        .PlotArea.Interior.Color = CLR_WHITE
        .ChartArea.Border.LineStyle = xlContinuous
        .ChartArea.Border.Color = CLR_GRAY_LIGHT
        .ChartArea.Interior.Color = CLR_WHITE
        .HasLegend = True
        .Legend.Font.Name = "Segoe UI"
        .Legend.Font.size = 8
        .Legend.Position = xlLegendPositionBottom
    End With

    ' CHART 3 — Age Distribution (horizontal bar)
    Dim co3 As ChartObject
    Set co3 = ws.ChartObjects.Add(ws.Cells(r, COL_S + 15).Left, baseTop, 280, 210)

    With co3.Chart
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.text = "Age Group Distribution"
        .ChartTitle.Font.Name = "Segoe UI Semibold"
        .ChartTitle.Font.size = 11
        .ChartTitle.Font.bold = True
        .ChartTitle.Font.Color = CLR_TEXT_PRIMARY

        Dim agLbls(4) As String
        agLbls(0) = "0-17 (Minor)"
        agLbls(1) = "18-25 (Young Adult)"
        agLbls(2) = "26-35 (Adult)"
        agLbls(3) = "36-50 (Middle Age)"
        agLbls(4) = "51+ (Senior)"

        Dim sr3 As Series
        Set sr3 = .SeriesCollection.NewSeries
        sr3.Values = ageGroups
        sr3.XValues = agLbls
        sr3.Name = "Residents"
        sr3.Format.Fill.ForeColor.RGB = CLR_ACCENT_BLUE
        sr3.Format.Line.Visible = msoFalse

        .Axes(xlCategory).TickLabels.Font.Name = "Segoe UI"
        .Axes(xlCategory).TickLabels.Font.size = 8
        .Axes(xlValue).TickLabels.Font.Name = "Segoe UI"
        .Axes(xlValue).TickLabels.Font.size = 8
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = CLR_GRAY_LIGHT
        .PlotArea.Interior.Color = CLR_WHITE
        .ChartArea.Border.LineStyle = xlContinuous
        .ChartArea.Border.Color = CLR_GRAY_LIGHT
        .ChartArea.Interior.Color = CLR_WHITE
        .HasLegend = False
    End With
End Sub

' =========================================================
' FOOTER
' =========================================================
Private Sub DrawFooter(ws As Worksheet, r As Long)
    With ws.Range(ws.Cells(r, COL_S), ws.Cells(r, COL_S + 19))
        .Merge
        .Value = "ERPAT System v3.0  •  Barangay Resident ID Management  •  © " & Year(Now) & " Municipality of Balayan, Batangas  •  Powered by Excel VBA"
        .Interior.Color = RGB(245, 245, 245)
        .Font.Name = "Segoe UI"
        .Font.size = 7
        .Font.Color = CLR_TEXT_SECONDARY
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = CLR_GRAY_MED
    End With
    ws.ROWS(r).RowHeight = 18
End Sub

' =========================================================
' ENHANCED UTILITIES
' =========================================================
Private Sub EnhancedSectionLabel(ws As Worksheet, r As Long, txt As String, accentClr As Long)
    With ws.Range(ws.Cells(r, COL_S), ws.Cells(r, COL_S + 19))
        .Merge
        .Value = "  " & txt
        .Interior.Color = BlendToWhite(accentClr, 0.92)
        .Font.Name = "Segoe UI Semibold"
        .Font.bold = True
        .Font.size = 10
        .Font.Color = accentClr
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = accentClr
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = CLR_GRAY_LIGHT
    End With
    ws.ROWS(r).RowHeight = 24
End Sub

Private Sub EnhancedMiniTableHeader(ws As Worksheet, r As Long, c As Long, _
                                     h1 As String, h2 As String, h3 As String, clr As Long)
    Dim hdrs(2) As String: hdrs(0) = h1: hdrs(1) = h2: hdrs(2) = h3
    Dim ww(2) As Long: ww(0) = 3: ww(1) = 1: ww(2) = 1
    Dim off As Long: off = 0
    Dim j As Long
    For j = 0 To 2
        With ws.Range(ws.Cells(r, c + off), ws.Cells(r, c + off + ww(j) - 1))
            .Merge
            .Value = hdrs(j)
            .Interior.Color = clr
            .Font.Name = "Segoe UI Semibold"
            .Font.Color = CLR_WHITE
            .Font.bold = True
            .Font.size = 9
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        off = off + ww(j)
    Next j
    ws.ROWS(r).RowHeight = 22
End Sub

Private Sub EnhancedMiniTableRow(ws As Worksheet, r As Long, c As Long, _
                                  v1 As String, v2 As String, v3 As String, isAlt As Boolean)
    Dim bg As Long
    bg = IIf(isAlt, RGB(248, 252, 250), CLR_WHITE)
    Dim ww(2) As Long: ww(0) = 3: ww(1) = 1: ww(2) = 1
    Dim vals(2) As String: vals(0) = v1: vals(1) = v2: vals(2) = v3
    Dim off As Long: off = 0
    Dim j As Long
    For j = 0 To 2
        With ws.Range(ws.Cells(r, c + off), ws.Cells(r, c + off + ww(j) - 1))
            .Merge
            .Value = vals(j)
            .Interior.Color = bg
            .Font.Name = "Segoe UI"
            .Font.size = 9
            .HorizontalAlignment = IIf(j = 0, xlLeft, xlCenter)
            .VerticalAlignment = xlCenter
            If j = 0 Then .IndentLevel = 1
            If j = 1 Then .Font.bold = True
        End With
        off = off + ww(j)
    Next j
    ws.ROWS(r).RowHeight = 18
End Sub

Private Sub EnhancedBoxBorder(ws As Worksheet, r1 As Long, c1 As Long, r2 As Long, c2 As Long)
    With ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = CLR_GRAY_MED
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = CLR_GRAY_MED
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = CLR_GRAY_MED
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = CLR_GRAY_MED
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = CLR_GRAY_LIGHT
        .Borders(xlInsideHorizontal).LineStyle = xlDot
        .Borders(xlInsideHorizontal).Color = CLR_GRAY_LIGHT
    End With
End Sub

Private Function SafeColIndex(tbl As ListObject, colName As String, fallback As Long) As Long
    On Error Resume Next
    Dim idx As Long
    idx = tbl.ListColumns(colName).Index
    On Error GoTo 0
    SafeColIndex = IIf(idx > 0, idx, fallback)
End Function

Private Function BlendToWhite(clr As Long, pct As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = (clr Mod 256)
    g = ((clr \ 256) Mod 256)
    b = ((clr \ 65536) Mod 256)
    r = CLng(r + (255 - r) * pct)
    g = CLng(g + (255 - g) * pct)
    b = CLng(b + (255 - b) * pct)
    BlendToWhite = RGB(r, g, b)
End Function

' =========================================================
' AUTO-REFRESH EVENT — PASTE INTO Sheet1 CODE WINDOW
' =========================================================
'
' Private Sub Worksheet_Change(ByVal Target As Range)
'     Dim tbl As ListObject
'     On Error Resume Next
'     Set tbl = Me.ListObjects("tblResidents")
'     On Error GoTo 0
'
'     If tbl Is Nothing Then Exit Sub
'     If tbl.DataBodyRange Is Nothing Then Exit Sub
'     If Intersect(Target, tbl.DataBodyRange) Is Nothing Then Exit Sub
'     If Target.Cells.Count > 50 Then Exit Sub
'
'     Application.EnableEvents = False
'     BuildERPATDashboard
'     Application.EnableEvents = True
' End Sub


