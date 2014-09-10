Attribute VB_Name = "AHF_File"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc  : Function FileExists
' Date  : 10/10/2012
' Type  : Boolean
' Desc  : Checks if a file exists and can be read
' Ex    : FileExists "C:\autoexec.bat"
'---------------------------------------------------------------------------------------
Function FileExists(ByVal FilePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Remove trailing backslash
    If InStr(Len(FilePath), FilePath, "\") > 0 Then
        FilePath = Left(FilePath, Len(FilePath) - 1)
    End If

    'Check to see if the file exists and has read access
    On Error GoTo File_Error
    If fso.FileExists(FilePath) Then
        fso.OpenTextFile(FilePath, 1).Read 0
        FileExists = True
    Else
        FileExists = False
    End If
    On Error GoTo 0

    Exit Function

File_Error:
    FileExists = False
End Function

'---------------------------------------------------------------------------------------
' Proc  : Function FolderExists
' Date  : 10/10/2012
' Type  : Boolean
' Desc  : Checks if a folder exists
' Ex    : FolderExists "C:\Program Files\"
'---------------------------------------------------------------------------------------
Function FolderExists(ByVal sPath As String) As Boolean
    'Add trailing backslash
    If InStr(Len(sPath), sPath, "\") = 0 Then sPath = sPath & "\"
    'If the folder exists return true
    On Error GoTo File_Error
    If Dir(sPath, vbDirectory) <> "" Then FolderExists = True
    On Error GoTo 0
    Exit Function

File_Error:
    FolderExists = False
End Function

'---------------------------------------------------------------------------------------
' Proc  : Sub RecMkDir
' Date  : 10/10/2012
' Desc  : Creates an entire directory tree
' Ex    : RecMkDir "C:\Dir1\Dir2\Dir3\"
'---------------------------------------------------------------------------------------
Sub RecMkDir(ByVal sPath As String)
    Dim sDirArray() As String   'Folder names
    Dim sDrive As String        'Base drive
    Dim sNewPath As String      'Path builder
    Dim LoopStart As Long       'Loop start number
    Dim i As Long               'Counter

    'Add trailing slash
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If

    'Split at each \
    If Left(sPath, 2) <> "\\" Then
        sDirArray = Split(sPath, "\")
        sDrive = sDirArray(0) & "\"
    Else
        sDirArray = Split(sPath, "\")
        sDrive = "\\" & sDirArray(2) & "\"
    End If

    'Determine where in the array to start the loop
    If sDrive = "\\" & sDirArray(2) & "\" Then
        If Not FolderExists(sDrive & sDirArray(3)) Then
            Err.Raise 76, "RecMkDir", "Network share not found."
        End If

        LoopStart = 3
    Else
        LoopStart = 1
    End If

    'Loop through each directory
    For i = LoopStart To UBound(sDirArray) - 1
        If Len(sNewPath) = 0 Then
            sNewPath = sDrive & sNewPath & sDirArray(i) & "\"
        Else
            sNewPath = sNewPath & sDirArray(i) & "\"
        End If

        If Not FolderExists(sNewPath) And Len(sNewPath) > 3 Then
            MkDir sNewPath
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : DeleteFile
' Date : 3/19/2013
' Desc : Deletes a file
'---------------------------------------------------------------------------------------
Sub DeleteFile(FileName As String)
    On Error Resume Next
    Kill FileName
End Sub
