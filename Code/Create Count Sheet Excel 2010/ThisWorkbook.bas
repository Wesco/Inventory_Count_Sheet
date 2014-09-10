VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Activate
    Worksheets("Macro").Select
    CleanUp
End Sub

Private Sub Workbook_Open()
    Range("G12").Formula = "=HYPERLINK(" & """" & ThisWorkbook.Path & """" & "," & """" & ThisWorkbook.Path & """" & ")"
End Sub