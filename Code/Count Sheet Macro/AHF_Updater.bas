Attribute VB_Name = "AHF_Updater"
Option Explicit

Private Enum Ver
    Major
    Minor
    Patch
End Enum

Private Declare Function ShellExecute _
                          Lib "shell32.dll" Alias "ShellExecuteA" ( _
                              ByVal hWnd As Long, _
                              ByVal Operation As String, _
                              ByVal FileName As String, _
                              Optional ByVal Parameters As String, _
                              Optional ByVal Directory As String, _
                              Optional ByVal WindowStyle As Long = vbMaximizedFocus _
                            ) As Long

'---------------------------------------------------------------------------------------
' Proc : IncrementMajor
' Date : 9/4/2013
' Desc : Increments the macros major version number (major.minor.patch)
'---------------------------------------------------------------------------------------
Sub IncrementMajor()
    IncrementVer Major
End Sub

'---------------------------------------------------------------------------------------
' Proc : IncrementMinorVersion
' Date : 4/24/2013
' Desc : Increments the macros minor version number (major.minor.patch)
'---------------------------------------------------------------------------------------
Sub IncrementMinor()
    IncrementVer Minor
End Sub

'---------------------------------------------------------------------------------------
' Proc : IncrementPatch
' Date : 9/4/2013
' Desc : Increments the macros patch number (major.minor.patch)
'---------------------------------------------------------------------------------------
Sub IncrementPatch()
    IncrementVer Patch
End Sub

'---------------------------------------------------------------------------------------
' Proc : IncrementVer
' Date : 9/4/2013
' Desc :
'---------------------------------------------------------------------------------------
Private Sub IncrementVer(Version As Ver)
    Dim Path As String
    Dim Ver As Variant
    Dim FileNum As Integer
    Dim i As Integer

    Path = Left(ThisWorkbook.fullName, InStr(1, ThisWorkbook.fullName, ThisWorkbook.Name, vbTextCompare) - 1) & "Version.txt"
    FileNum = FreeFile

    If FileExists(Path) = True Then
        Open Path For Input As #FileNum
        Line Input #FileNum, Ver
        Close FileNum

        'Split version number
        Ver = Split(Ver, ".")

        'Increment version
        Select Case Version
            Case Major
                Ver(0) = CInt(Ver(0)) + 1
            Case Minor
                Ver(1) = CInt(Ver(1)) + 1
            Case Patch
                Ver(2) = CInt(Ver(2)) + 1
        End Select

        'Combine version
        Ver = Ver(0) & "." & Ver(1) & "." & Ver(2)

        Open Path For Output As #FileNum
        Print #FileNum, Ver
        Close #FileNum
    Else
        Open Path For Output As #FileNum
        Print #FileNum, "1.0.0"
        Close #FileNum
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : CheckForUpdates
' Date : 4/24/2013
' Desc : Checks to see if the macro is up to date
'---------------------------------------------------------------------------------------
Sub CheckForUpdates(RepoName As String, LocalVer As String)
    Dim RemoteVer As Variant
    Dim RegEx As Variant
    Dim Result As Integer

    On Error GoTo UPDATE_ERROR
    Set RegEx = CreateObject("VBScript.RegExp")

    'Try to get the contents of the text file
    RemoteVer = DownloadTextFile("http://br3615gaps.wescodist.com/" & RepoName & "/Version.txt")
    RemoteVer = Replace(RemoteVer, vbLf, "")
    RemoteVer = Replace(RemoteVer, vbCr, "")

    'Expression to verify the data retrieved is a version number
    RegEx.Pattern = "^[0-9]+\.[0-9]+\.[0-9]+$"

    If RegEx.Test(RemoteVer) Then
        If Not RemoteVer = LocalVer Then
            Result = MsgBox("An update is available. Would you like to download the latest version now?", vbYesNo, "Update Available")
            If Result = vbYes Then
                'Opens github release page in the default browser, maximised with focus by default
                ShellExecute 0, "Open", "http://github.com/Wesco/" & RepoName & "/releases/"
                ThisWorkbook.Saved = True
                If Workbooks.Count = 1 Then
                    Application.Quit
                Else
                    ThisWorkbook.Close
                End If
            End If
        End If
    Else
        GoTo UPDATE_ERROR
    End If
    On Error GoTo 0
    Exit Sub

UPDATE_ERROR:
    If MsgBox("An error occured while checking for updates." & vbCrLf & vbCrLf & _
              "Would you like to open the website to download the latest version?", vbYesNo, _
              "Version " & LocalVer) = vbYes Then
        ShellExecute 0, "Open", "http://github.com/Wesco/" & RepoName & "/releases/"
        If Workbooks.Count = 1 Then
            Application.Quit
        Else
            ThisWorkbook.Close
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : DownloadTextFile
' Date : 4/25/2013
' Desc : Returns the contents of a text file from a website
'---------------------------------------------------------------------------------------
Private Function DownloadTextFile(URL As String) As String
    Dim success As Boolean
    Dim responseText As String
    Dim oHTTP As Variant

    Set oHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

    oHTTP.Open "GET", URL, False
    oHTTP.Send
    success = oHTTP.WaitForResponse()

    If Not success Then
        DownloadTextFile = ""
        Exit Function
    End If

    responseText = oHTTP.responseText
    Set oHTTP = Nothing

    DownloadTextFile = responseText
End Function
