Attribute VB_Name = "modDashboardProtection"
Option Explicit

Public Sub SafeRefreshAllPivots()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pc As PivotCache

    On Error GoTo handler

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With

    '========================================================
    ' REFRESH ALL PIVOT CACHES
    '========================================================
    For Each pc In ThisWorkbook.PivotCaches

        On Error Resume Next
        pc.Refresh
        On Error GoTo handler

    Next pc

    '========================================================
    ' REPAIR EMPTY FILTER STATES
    '========================================================
    For Each ws In ThisWorkbook.Worksheets

        For Each pt In ws.PivotTables

            On Error Resume Next

            pt.ManualUpdate = True

            ' Clear broken filters
            pt.ClearAllFilters
            
            Dim pf As PivotField

For Each pf In pt.PivotFields

    On Error Resume Next

    pf.EnableMultiplePageItems = True

    On Error GoTo handler

Next pf

            ' Force refresh
            pt.RefreshTable

            pt.ManualUpdate = False

            On Error GoTo handler

        Next pt

    Next ws

CleanExit:

    With Application
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Exit Sub

handler:

    Resume CleanExit

End Sub
