VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents DayLabel As MSForms.Label
Attribute DayLabel.VB_VarHelpID = -1
Public Index As Long
Public ParentCal As Object  ' frmCalendar

Private Sub DayLabel_Click()
    ParentCal.DayClick Index
End Sub
