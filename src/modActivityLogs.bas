Attribute VB_Name = "modActivityLogs"
Option Explicit

Public Sub WriteSystemLog( _
    ByVal ActionType As String, _
    ByVal ResidentID As String, _
    ByVal FullName As String)

    Dim logFolder As String
    Dim logFile As String
    Dim fileNum As Integer
    Dim logLine As String

    On Error GoTo handler

    '========================================================
    ' CREATE LOGS DIRECTORY
    '========================================================
    logFolder = ThisWorkbook.Path & "\Logs"

    If Dir(logFolder, vbDirectory) = "" Then
        MkDir logFolder
    End If

    '========================================================
    ' DAILY LOG FILE
    '========================================================
    logFile = logFolder & "\LOG_" & Format(Date, "yyyy-mm-dd") & ".txt"

    '========================================================
    ' BUILD LOG ENTRY
    '========================================================
    logLine = "[" & Format(Time, "hh:mm:ss AM/PM") & "] " & _
              "USER: " & Environ("Username") & _
              " | ACTION: " & UCase(ActionType) & _
              " | ID: " & ResidentID & _
              " | NAME: " & FullName

    '========================================================
    ' WRITE TO FILE
    '========================================================
    fileNum = FreeFile

    Open logFile For Append As #fileNum
        Print #fileNum, logLine
    Close #fileNum

    Exit Sub

handler:

    MsgBox "Log write failed: " & Err.Description, vbExclamation

End Sub

