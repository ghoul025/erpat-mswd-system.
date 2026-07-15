VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmCalendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPick As Variant
Private mHandlers As Collection ' clsCalDay handlers

' ===== PUBLIC API =====
Public Function PickDate(Optional ByVal startDate As Variant) As Variant
    Dim d As Date
    If IsDate(startDate) Then
        d = CDate(startDate)
    Else
        d = Date
    End If

    mPick = Empty

    BuildDayGrid
    BuildMonthYearDropdowns d
    RenderCalendar

    Me.Show vbModal
    PickDate = mPick
End Function

' ===== UI BUTTONS =====
Private Sub cmdPrev_Click()
    ' Previous month via dropdown
    If cboMonth.ListIndex > 0 Then
        cboMonth.ListIndex = cboMonth.ListIndex - 1
    Else
        ' wrap to Dec previous year
        If cboYear.ListIndex > 0 Then
            cboYear.ListIndex = cboYear.ListIndex - 1
            cboMonth.ListIndex = 11
        End If
    End If
    ' Render handled by Change event
End Sub

Private Sub cmdNext_Click()
    ' Next month via dropdown
    If cboMonth.ListIndex < 11 Then
        cboMonth.ListIndex = cboMonth.ListIndex + 1
    Else
        ' wrap to Jan next year
        If cboYear.ListIndex < cboYear.ListCount - 1 Then
            cboYear.ListIndex = cboYear.ListIndex + 1
            cboMonth.ListIndex = 0
        End If
    End If
    ' Render handled by Change event
End Sub

Private Sub cmdToday_Click()
    mPick = Date
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    mPick = Empty
    Me.Hide
End Sub

' ===== DROPDOWN EVENTS =====
Private Sub cboMonth_Change()
    If cboMonth.ListIndex >= 0 And cboYear.ListIndex >= 0 Then RenderCalendar
End Sub

Private Sub cboYear_Change()
    If cboMonth.ListIndex >= 0 And cboYear.ListIndex >= 0 Then RenderCalendar
End Sub

' ===== MONTH/YEAR DROPDOWNS =====
Private Sub BuildMonthYearDropdowns(ByVal baseDate As Date)
    Dim i As Long
    Dim y As Long, yStart As Long, yEnd As Long

    ' Build Month list (Jan..Dec)
    cboMonth.Clear
    For i = 1 To 12
        cboMonth.AddItem Format$(DateSerial(2000, i, 1), "mmmm")
    Next i

    ' Build Year list (adjust range as you want)
    yStart = Year(Date) - 70
    yEnd = Year(Date)

    cboYear.Clear
    For y = yStart To yEnd
        cboYear.AddItem CStr(y)
    Next y

    ' Set selected values from baseDate
    cboMonth.ListIndex = Month(baseDate) - 1
    cboYear.Value = CStr(Year(baseDate))
End Sub

Private Function GetBaseMonthFirstDay() As Date
    ' cboMonth is 0-based (0=Jan)
    Dim m As Long, y As Long
    m = cboMonth.ListIndex + 1
    y = CLng(cboYear.Value)
    GetBaseMonthFirstDay = DateSerial(y, m, 1)
End Function

' ===== GRID GENERATION (AUTO CREATE LABELS) =====
Private Sub BuildDayGrid()
    Dim i As Long
    Dim lbl As MSForms.Label
    Dim h As clsCalDay

    ' Already built?
    On Error Resume Next
    Set lbl = Me.Controls("d1")
    On Error GoTo 0
    If Not lbl Is Nothing Then Exit Sub

    Set mHandlers = New Collection

    ' layout settings (tweak if you want)
    Dim startLeft As Single: startLeft = 12
    Dim startTop As Single:  startTop = 70
    Dim w As Single: w = 32
    Dim hgt As Single: hgt = 22
    Dim gap As Single: gap = 4

    For i = 1 To 42
        Set lbl = Me.Controls.Add("Forms.Label.1", "d" & i, True)

        With lbl
            .Caption = ""
            .Tag = ""
            .TextAlign = fmTextAlignCenter
            .BorderStyle = fmBorderStyleSingle
            .SpecialEffect = fmSpecialEffectSunken
            .BackStyle = fmBackStyleOpaque
            .Font.Name = "Calibri"
            .Font.size = 10
            .Width = w
            .Height = hgt
            .Left = startLeft + ((i - 1) Mod 7) * (w + gap)
            .Top = startTop + Int((i - 1) / 7) * (hgt + gap)
        End With

        Set h = New clsCalDay
        Set h.DayLabel = lbl
        h.Index = i
        Set h.ParentCal = Me
        mHandlers.Add h
    Next i
End Sub

' ===== CALENDAR RENDER =====
Private Sub RenderCalendar()
    Dim firstDay As Date, startCell As Long, daysInMonth As Long
    Dim i As Long, d As Long

    firstDay = GetBaseMonthFirstDay()

    ' Optional label title (if you keep lblMonth)
    On Error Resume Next
    lblMonth.Caption = Format$(firstDay, "mmmm yyyy")
    On Error GoTo 0

    ' Clear all cells
    For i = 1 To 42
        Me.Controls("d" & i).Caption = ""
        Me.Controls("d" & i).Tag = ""
    Next i

    ' Monday-first alignment
    startCell = Weekday(firstDay, vbMonday)
    daysInMonth = Day(DateSerial(Year(firstDay), Month(firstDay) + 1, 0))

    d = 1
    For i = startCell To startCell + daysInMonth - 1
        Me.Controls("d" & i).Caption = CStr(d)
        Me.Controls("d" & i).Tag = CStr(DateSerial(Year(firstDay), Month(firstDay), d))
        d = d + 1
    Next i
End Sub

' ===== CALLED BY clsCalDay click =====
Public Sub DayClick(ByVal idx As Long)
    Dim t As String
    t = CStr(Me.Controls("d" & idx).Tag)
    If t <> "" Then
        mPick = CDate(t)
        Me.Hide
    End If
End Sub

Private Sub UserForm_Click()

End Sub
