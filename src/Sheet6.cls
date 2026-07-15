VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    ApplySystemLock
End Sub

Public Sub ApplySystemLock()

    On Error Resume Next

    '====================================================
    ' VISUAL LOCK
    '====================================================
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayGridlines = False
        .Zoom = 130
    End With

    '====================================================
    ' HARD SCROLL LOCK
    '====================================================
    Me.ScrollArea = "A1:M50"

    '====================================================
    ' FORCE USER BACK TO TOP LEFT
    '====================================================
    Application.GoTo Me.Range("A1"), True

    On Error GoTo 0

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Intersect(Target, Me.Range("A1:M50")) Is Nothing Then
        Application.GoTo Me.Range("A1"), True
    End If

End Sub
