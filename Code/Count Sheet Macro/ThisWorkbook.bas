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

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Environ("username") <> "TReische" Then
        Cancel = True
    Else
        ExportCode
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Saved = True
End Sub

Private Sub Workbook_Open()
    Range("G12").Formula = "=HYPERLINK(" & """" & ThisWorkbook.Path & """" & "," & """" & ThisWorkbook.Path & """" & ")"
    CheckForUpdates RepositoryName, VersionNumber
End Sub
