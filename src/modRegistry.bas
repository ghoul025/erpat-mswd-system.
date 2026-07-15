Attribute VB_Name = "modRegistry"
Option Explicit

Public Const REG_SHEET As String = "Sheet1"      ' change if needed
Public Const REG_TABLE As String = "tblResidents"
Public Const BRGY_RANGE As String = "rngBarangays"

Public Sub ShowResidentForm()
    frmResident.Show vbModeless
End Sub

Public Function GetRegistryTable() As ListObject
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(REG_SHEET)
    Set GetRegistryTable = ws.ListObjects(REG_TABLE)
End Function

Public Function NextID() As String
    Dim tbl As ListObject
    Set tbl = GetRegistryTable()

    Dim curYear As Long
    curYear = Year(Date)

    Dim nextSeq As Long
    nextSeq = NextSequenceForYear(tbl, curYear)

    NextID = CStr(curYear) & "-" & Format$(nextSeq, "00")   ' 2026-01
End Function

Private Function NextSequenceForYear(ByVal tbl As ListObject, ByVal y As Long) As Long
    Dim maxSeq As Long: maxSeq = 0

    If tbl.ListRows.Count = 0 Then
        NextSequenceForYear = 1
        Exit Function
    End If

    Dim rng As Range, cell As Range, s As String
    Dim parts() As String
    Dim idYear As Long, idSeq As Long

    Set rng = tbl.ListColumns(1).DataBodyRange ' ID column

    For Each cell In rng.Cells
        s = Trim$(CStr(cell.Value))
        If Len(s) > 0 Then
            parts = Split(s, "-")
            If UBound(parts) = 1 Then
                If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
                    idYear = CLng(parts(0))
                    idSeq = CLng(parts(1))
                    If idYear = y Then
                        If idSeq > maxSeq Then maxSeq = idSeq
                    End If
                End If
            End If
        End If
    Next cell

    NextSequenceForYear = maxSeq + 1
End Function

Public Function IDExists(ByVal idVal As String) As Boolean
    Dim tbl As ListObject
    Set tbl = GetRegistryTable()

    IDExists = False
    If tbl.ListRows.Count = 0 Then Exit Function

    Dim f As Range
    Set f = tbl.ListColumns(1).DataBodyRange.Find(What:=Trim$(idVal), LookIn:=xlValues, LookAt:=xlWhole)

    IDExists = Not (f Is Nothing)
End Function

Public Function NormalizePhone(ByVal s As String) As String
    ' remove spaces/dashes/parentheses
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        If (ch Like "[0-9]") Or ch = "+" Then out = out & ch
    Next i
    NormalizePhone = out
End Function

Public Function IsValidPHMobile(ByVal raw As String) As Boolean
    Dim s As String
    s = NormalizePhone(raw)

    ' Accept:
    ' 09XXXXXXXXX  (11 digits)
    ' +63XXXXXXXXXX (13 chars incl +)
    ' 63XXXXXXXXXX (12 digits)  -> we can allow too, but keep strict if you want
    If Len(s) = 11 Then
        IsValidPHMobile = (Left$(s, 2) = "09") And (s Like "0##########")
    ElseIf Len(s) = 13 Then
        IsValidPHMobile = (Left$(s, 3) = "+63") And (mid$(s, 4) Like "##########")
    ElseIf Len(s) = 12 Then
        IsValidPHMobile = (Left$(s, 2) = "63") And (mid$(s, 3) Like "##########")
    Else
        IsValidPHMobile = False
    End If
End Function

Public Function ToProperDate(ByVal s As String) As Variant
    ' tries to convert, returns Empty if invalid
    On Error GoTo bad
    If Trim$(s) = "" Then
        ToProperDate = Empty
        Exit Function
    End If
    ToProperDate = CDate(s)
    Exit Function
bad:
    ToProperDate = Empty
End Function
