Attribute VB_Name = "modBatchPrint8Up"
Option Explicit

' ========= SOURCE =========
Private Const SRC_SHEET As String = "Sheet1"
Private Const SRC_TABLE As String = "tblResidents"

Private Const PRINT_FRONT_SHEET As String = "PRINT_FRONT"
Private Const PRINT_BACK_SHEET As String = "PRINT_BACK"

' ========= LAYOUT =========
Private Const COLS As Long = 2
Private Const ROWS As Long = 4
Private Const PER_PAGE As Long = 8

Private Const GAP_X As Double = 6
Private Const GAP_Y As Double = 6

' ========= PAGE SIZE =========
Private Const PAGE_H_IN As Double = 11      ' Letter paper height (inches)
Private Const PAGE_W_IN As Double = 8.5     ' Letter paper width (inches)

' ========= SAFE START =========
Private Const SAFE_X As Double = 0.2
Private Const SAFE_Y As Double = 0.2
Private Const BACK_CALIB_X As Double = 0.17  ' inches
Private Const BACK_CALIB_Y As Double = 0     ' inches

' ========= CARD SIZE =========
Private Const CARD_W_IN As Double = 3.4
Private Const CARD_H_IN As Double = 2.25

' ========= TEMPLATES =========
Private Const TEMPLATE_FRONT As String = "ERPAT ID.png"
Private Const TEMPLATE_BACK As String = "erpat_back.png"

' ========= FRONT TEXT POS =========
Private Const X_NAME As Double = 0.42
Private Const Y_NAME As Double = 0.22

Private Const X_ADDR As Double = 0.47
Private Const Y_ADDR As Double = 0.38

Private Const X_DOB As Double = 0.41
Private Const Y_DOB As Double = 0.49

Private Const X_ISSUED As Double = 0.66
Private Const Y_ISSUED As Double = 0.49

Private Const X_GENDER As Double = 0.87
Private Const Y_GENDER As Double = 0.49

Private Const X_ID As Double = 0.1
Private Const Y_ID As Double = 0.65

' ========= BACK TEXT POS =========
Private Const X_EMERGENCY_NAME As Double = 0.25
Private Const Y_EMERGENCY_NAME As Double = 0.13

Private Const X_EMERGENCY_ADDRESS As Double = 0.3
Private Const Y_EMERGENCY_ADDRESS As Double = 0.28

Private Const X_EMERGENCY_CONTACT As Double = 0.34
Private Const Y_EMERGENCY_CONTACT As Double = 0.4

' ========= FONT SIZES =========
Private Const FS_NAME As Double = 11
Private Const FS_ADDR As Double = 11
Private Const FS_SMALL As Double = 9
Private Const FS_ID As Double = 12

Private Const FS_BACK_NAME As Double = 11
Private Const FS_BACK_ADDR As Double = 9
Private Const FS_BACK_CONTACT As Double = 10

' ========= GENDER COLUMN =========
Private Const COL_GENDER As Long = 6   ' <-- Adjust this to your actual Gender column index

' =========================================================
' MAIN
' =========================================================
Public Sub RenderSelectedIDs_FrontBack_SeparateSheets()

    On Error GoTo Fail

    Dim wsSrc As Worksheet
    Dim wsFront As Worksheet
    Dim wsBack As Worksheet

    Dim tbl As ListObject
    Dim rowsToPrint As Collection

    Dim tplFront As String
    Dim tplBack As String

    Dim cardW As Double
    Dim cardH As Double

    Dim i As Long

    Set wsSrc = ThisWorkbook.Worksheets(SRC_SHEET)
    Set wsFront = ThisWorkbook.Worksheets(PRINT_FRONT_SHEET)
    Set wsBack = ThisWorkbook.Worksheets(PRINT_BACK_SHEET)

    Set tbl = wsSrc.ListObjects(SRC_TABLE)

    If tbl.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Selection, tbl.DataBodyRange) Is Nothing Then Exit Sub

    tplFront = GetTemplatePath(TEMPLATE_FRONT)
    tplBack = GetTemplatePath(TEMPLATE_BACK)

    If tplFront = "" Or tplBack = "" Then Exit Sub

    Set rowsToPrint = GetSelectedTableRowIndexes(tbl, Selection)

    If rowsToPrint.Count = 0 Then
        MsgBox "No valid rows selected.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    PreparePrintSheet wsFront
    PreparePrintSheet wsBack

    cardW = Application.InchesToPoints(CARD_W_IN)
    cardH = Application.InchesToPoints(CARD_H_IN)

    ' FRONT
    For i = 1 To rowsToPrint.Count
        RenderOneCard wsFront, tbl, rowsToPrint(i), i, tplFront, cardW, cardH, True
    Next i

    ' BACK
    For i = 1 To rowsToPrint.Count
        RenderOneCard wsBack, tbl, rowsToPrint(i), i, tplBack, cardW, cardH, False
    Next i

    Application.ScreenUpdating = True

    wsFront.Activate

    MsgBox "DUAL-SIDE RENDER COMPLETE " & Chr(10) & rowsToPrint.Count & " card(s) generated across " & _
           WorksheetFunction.Ceiling_Math(rowsToPrint.Count / PER_PAGE, 1) & " page(s).", vbInformation

    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "ERROR: " & Err.Description, vbCritical

End Sub

' =========================================================
' RENDER ONE CARD
' =========================================================
Private Sub RenderOneCard(ws As Worksheet, _
                          tbl As ListObject, _
                          rowIndex As Long, _
                          n As Long, _
                          tpl As String, _
                          cardW As Double, _
                          cardH As Double, _
                          isFront As Boolean)

    Dim r As Long
    Dim c As Long
    Dim pageNum As Long
    Dim posOnPage As Long
    Dim mirrorC As Long

    Dim leftPt As Double
    Dim topPt As Double

    Dim startX As Double
    Dim startY As Double
    Dim pageOffsetY As Double

    ' =========================
    ' MULTI-PAGE POSITION
    ' =========================
    pageNum = (n - 1) \ PER_PAGE            ' 0-based page index
    posOnPage = (n - 1) Mod PER_PAGE        ' position within current page

    r = posOnPage \ COLS
    c = posOnPage Mod COLS

    ' =========================
    ' MIRROR LOGIC (back side)
    ' =========================
    mirrorC = c

    If Not isFront Then
        mirrorC = (COLS - 1) - c
    End If

    ' =========================
    ' ORIGIN + PAGE OFFSET
    ' =========================
    startX = Application.InchesToPoints(SAFE_X)
    startY = Application.InchesToPoints(SAFE_Y)

    pageOffsetY = Application.InchesToPoints(PAGE_H_IN) * pageNum

    leftPt = startX + mirrorC * (cardW + GAP_X)
    topPt = startY + r * (cardH + GAP_Y) + pageOffsetY

    ' =========================
    ' BACK CALIBRATION
    ' =========================
    If Not isFront Then
        leftPt = leftPt + Application.InchesToPoints(BACK_CALIB_X)
        topPt = topPt + Application.InchesToPoints(BACK_CALIB_Y)
    End If

    ws.Shapes.AddPicture tpl, msoFalse, msoTrue, leftPt, topPt, cardW, cardH

    If isFront Then
        AddText ws, tbl, rowIndex, leftPt, topPt, cardW, cardH
    Else
        AddBackText ws, tbl, rowIndex, leftPt, topPt, cardW, cardH
    End If

End Sub

' =========================================================
' FRONT TEXT
' =========================================================
Private Sub AddText(ws As Worksheet, _
                    tbl As ListObject, _
                    i As Long, _
                    l As Double, _
                    t As Double, _
                    w As Double, _
                    h As Double)

    Dim firstName As String
    Dim middleName As String
    Dim lastName As String

    Dim middleInitial As String
    Dim FullName As String

    Dim residentAddress As String
    Dim gender As String

    firstName = Trim(tbl.DataBodyRange.Cells(i, 3).Value)
    middleName = Trim(tbl.DataBodyRange.Cells(i, 4).Value)
    lastName = Trim(tbl.DataBodyRange.Cells(i, 2).Value)
    residentAddress = Trim(tbl.DataBodyRange.Cells(i, 8).Value)
    gender = Trim(tbl.DataBodyRange.Cells(i, COL_GENDER).Value)

    ' =====================================================
    ' AUTO ADD "Brgy." IF MISSING
    ' =====================================================
    If LCase(Left(residentAddress, 5)) <> "brgy." Then
        residentAddress = "Brgy. " & residentAddress
    End If

    residentAddress = residentAddress & ", Balayan, Batangas"

    ' =====================================================
    ' MIDDLE INITIAL
    ' =====================================================
    If Len(middleName) > 0 Then
        middleInitial = Left(middleName, 1) & "."
    Else
        middleInitial = ""
    End If

    FullName = Trim(firstName & " " & middleInitial & " " & lastName)

    ' =====================================================
    ' RENDER FRONT FIELDS
    ' =====================================================
    AddBox ws, FullName, _
        l + w * X_NAME, _
        t + h * Y_NAME, _
        FS_NAME, True

    AddBox ws, residentAddress, _
        l + w * X_ADDR, _
        t + h * Y_ADDR, _
        FS_ADDR, False, 120, 32

    AddBox ws, FormatDateSafe(tbl.DataBodyRange.Cells(i, 5).Value), _
        l + w * X_DOB, _
        t + h * Y_DOB, _
        FS_SMALL, False

    AddBox ws, FormatDateSafe(tbl.DataBodyRange.Cells(i, 13).Value), _
        l + w * X_ISSUED, _
        t + h * Y_ISSUED, _
        FS_SMALL, False

    ' =====================================================
    ' GENDER — now reads from table (was hardcoded "M")
    ' =====================================================
    AddBox ws, "M", _
        l + w * X_GENDER, _
        t + h * Y_GENDER, _
        FS_SMALL, True
    AddBox ws, tbl.DataBodyRange.Cells(i, 1).Value, _
        l + w * X_ID, _
        t + h * Y_ID, _
        FS_ID, True

End Sub

' =========================================================
' BACK TEXT
' =========================================================
Private Sub AddBackText(ws As Worksheet, _
                        tbl As ListObject, _
                        i As Long, _
                        l As Double, _
                        t As Double, _
                        w As Double, _
                        h As Double)

    Dim colEmergencyPerson As Long
    Dim colEmergencyAddress As Long
    Dim colEmergencyContact As Long

    Dim emergencyAddress As String

    ' =====================================================
    ' SEARCH COLUMNS BY HEADER NAME
    ' =====================================================
    colEmergencyPerson = tbl.ListColumns("EMERGENCY CONTACT PERSON").Index
    colEmergencyAddress = tbl.ListColumns("ADDRESS").Index
    colEmergencyContact = tbl.ListColumns("EMERGENCY CONTACT NO.").Index

    emergencyAddress = Trim(tbl.DataBodyRange.Cells(i, colEmergencyAddress).Value)

    ' =====================================================
    ' AUTO ADD "Brgy." IF MISSING
    ' =====================================================
    If LCase(Left(emergencyAddress, 5)) <> "brgy." Then
        emergencyAddress = "Brgy. " & emergencyAddress
    End If

    emergencyAddress = emergencyAddress & ", Balayan, Batangas"

    ' =====================================================
    ' RENDER BACK FIELDS
    ' =====================================================
    AddBox ws, Trim(tbl.DataBodyRange.Cells(i, colEmergencyPerson).Value), _
        l + w * X_EMERGENCY_NAME, _
        t + h * Y_EMERGENCY_NAME, _
        FS_BACK_NAME, True

    AddBox ws, emergencyAddress, _
        l + w * X_EMERGENCY_ADDRESS, _
        t + h * Y_EMERGENCY_ADDRESS, _
        FS_BACK_ADDR, False, 90, 32

    AddBox ws, Trim(tbl.DataBodyRange.Cells(i, colEmergencyContact).Value), _
        l + w * X_EMERGENCY_CONTACT, _
        t + h * Y_EMERGENCY_CONTACT, _
        FS_BACK_CONTACT, True

End Sub

' =========================================================
' TEXTBOX CREATOR
' =========================================================
Private Sub AddBox(ws As Worksheet, _
                   txt As String, _
                   l As Double, _
                   t As Double, _
                   size As Double, _
                   bold As Boolean, _
                   Optional boxW As Double = 140, _
                   Optional boxH As Double = 40, _
                   Optional alignCenter As Boolean = False)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        l, t, boxW, boxH)

    With shp

        .TextFrame2.TextRange.text = txt
        .TextFrame2.TextRange.Font.size = size
        .TextFrame2.TextRange.Font.bold = bold

        ' =========================
        ' WORD WRAP + LIMITS
        ' =========================
        .TextFrame2.WordWrap = msoTrue
        .TextFrame2.AutoSize = msoAutoSizeNone

        ' =========================
        ' ALIGNMENT
        ' =========================
        If alignCenter Then
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        Else
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        End If

        ' =========================
        ' VERTICAL ALIGN
        ' =========================
        .TextFrame2.VerticalAnchor = msoAnchorMiddle

        ' =========================
        ' INTERNAL MARGINS
        ' =========================
        .TextFrame2.MarginLeft = 2
        .TextFrame2.MarginRight = 2
        .TextFrame2.MarginTop = 1
        .TextFrame2.MarginBottom = 1

        ' =========================
        ' HIDE BOX DESIGN
        ' =========================
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse

    End With

End Sub

' =========================================================
' HELPERS
' =========================================================
Private Function FormatDateSafe(v) As String

    If IsDate(v) Then
        FormatDateSafe = Format(v, "mm/dd/yyyy")
    Else
        FormatDateSafe = ""
    End If

End Function

Private Function GetTemplatePath(f As String) As String

    Dim p As String

    p = ThisWorkbook.Path & "\" & f

    If Dir(p) <> "" Then
        GetTemplatePath = p
    Else
        MsgBox f & " not found. Please ensure the file is in the same folder as this workbook.", vbCritical
        GetTemplatePath = ""
    End If

End Function

Private Function GetSelectedTableRowIndexes(tbl As ListObject, sel As Range) As Collection

    Dim c As New Collection
    Dim cell As Range
    Dim idx As Long
    Dim seen As Object

    Set seen = CreateObject("Scripting.Dictionary")

    For Each cell In sel
        idx = cell.Row - tbl.DataBodyRange.Row + 1
        If idx >= 1 And idx <= tbl.DataBodyRange.ROWS.Count Then
            If Not seen.Exists(idx) Then
                seen(idx) = True
                c.Add idx
            End If
        End If
    Next

    Set GetSelectedTableRowIndexes = c

End Function

Private Sub PreparePrintSheet(ws As Worksheet)

    ws.Cells.Clear

    Dim i As Long

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

End Sub

