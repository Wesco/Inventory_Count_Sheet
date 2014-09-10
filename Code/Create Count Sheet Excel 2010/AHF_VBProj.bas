Attribute VB_Name = "AHF_VBProj"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : ExportCode
' Date : 3/19/2013
' Desc : Exports all modules
'---------------------------------------------------------------------------------------
Sub ExportCode()
    Dim comp As Variant
    Dim codeFolder As String
    Dim FileName As String
    Dim File As String
    Dim WkbkPath As String


    'References Microsoft Visual Basic for Applications Extensibility 5.3
    AddReference "{0002E157-0000-0000-C000-000000000046}", 5, 3
    WkbkPath = Left$(ThisWorkbook.fullName, InStr(1, ThisWorkbook.fullName, ThisWorkbook.Name, vbTextCompare) - 1)
    codeFolder = WkbkPath & "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "\"

    On Error Resume Next
    If Dir(codeFolder) = "" Then
        RecMkDir codeFolder
    End If
    On Error GoTo 0

    'Remove all previously exported modules
    File = Dir(codeFolder)
    Do While File <> ""
        DeleteFile codeFolder & File
        File = Dir
    Loop

    'Export modules in current project
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1
                FileName = codeFolder & comp.Name & ".bas"
                comp.Export FileName
            Case 2
                FileName = codeFolder & comp.Name & ".cls"
                comp.Export FileName
            Case 3
                FileName = codeFolder & comp.Name & ".frm"
                comp.Export FileName
            Case 100
                If comp.Name = "ThisWorkbook" Then
                    FileName = codeFolder & comp.Name & ".bas"
                    comp.Export FileName
                End If
        End Select
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : AddReferences
' Date : 3/19/2013
' Desc : Adds a reference to VBProject
'---------------------------------------------------------------------------------------
Private Sub AddReference(GUID As String, Major As Integer, Minor As Integer)
    Dim ID As Variant
    Dim Ref As Variant
    Dim Result As Boolean


    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = GUID And Ref.Major = Major And Ref.Minor = Minor Then
            Result = True
            Exit For
        End If
    Next

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    If Result = False Then
        ThisWorkbook.VBProject.References.AddFromGuid GUID, Major, Minor
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : RemoveReferences
' Date : 3/19/2013
' Desc : Removes a reference from VBProject
'---------------------------------------------------------------------------------------
Sub RemoveReference(GUID As String, Major As Integer, Minor As Integer)
    Dim Ref As Variant


    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = GUID And Ref.Major = Major And Ref.Minor = Minor Then
            Application.VBE.ActiveVBProject.References.Remove Ref
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : ShowReferences
' Date : 4/4/2013
' Desc : Lists all VBProject references
'---------------------------------------------------------------------------------------
Sub ShowReferences()
    Dim i As Variant
    Dim n As Integer


    ThisWorkbook.Activate
    On Error GoTo SHEET_EXISTS
    Sheets("VBA References").Select
    ActiveSheet.Cells.Delete
    On Error GoTo 0

    [A1].Value = "Name"
    [B1].Value = "Description"
    [C1].Value = "GUID"
    [D1].Value = "Major"
    [E1].Value = "Minor"

    For i = 1 To ThisWorkbook.VBProject.References.Count
        n = i + 1
        With ThisWorkbook.VBProject.References(i)
            Cells(n, 1).Value = .Name
            Cells(n, 2).Value = .Description
            Cells(n, 3).Value = .GUID
            Cells(n, 4).Value = .Major
            Cells(n, 5).Value = .Minor
        End With
    Next
    Columns.AutoFit

    Exit Sub

SHEET_EXISTS:
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count), Count:=1
    ActiveSheet.Name = "VBA References"
    Resume Next
End Sub
