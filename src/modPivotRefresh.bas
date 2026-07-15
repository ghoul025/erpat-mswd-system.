Attribute VB_Name = "modPivotRefresh"
Option Explicit

Public Sub RefreshSystemAnalytics()

    Dim objCache As PivotCache
    Dim wksDash As Worksheet

    On Error GoTo CleanFail

    '========================================================
    ' PERFORMANCE LOCK
    '========================================================
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .StatusBar = "Refreshing dashboard analytics..."
    End With

    '========================================================
    ' SAFE PIVOT CACHE REFRESH LOOP
    '========================================================
    SafeRefreshAllPivots

    '========================================================
    ' REFRESH DASHBOARD WORKSHEET
    '========================================================
    On Error Resume Next

    Set wksDash = ThisWorkbook.Worksheets("DASHBOARD")

    If Not wksDash Is Nothing Then
        wksDash.Calculate
    End If

    On Error GoTo CleanFail

CleanExit:

    '========================================================
    ' RESTORE EXCEL STATE
    '========================================================
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .ScreenUpdating = True
        .StatusBar = False
    End With

    Exit Sub

CleanFail:

    MsgBox "Dashboard refresh error: " & Err.Description, vbExclamation

    Resume CleanExit

End Sub
