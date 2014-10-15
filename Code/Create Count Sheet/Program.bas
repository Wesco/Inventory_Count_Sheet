Attribute VB_Name = "Program"
Option Explicit

Public Const VersionNumber As String = "1.0.0"
Public Const RepositoryName As String = "Inventory_Count_Sheet"

Sub CountSheet()
    Dim i As Long
    Dim x As Long
    Dim Path As String
    Dim vArray As Variant
    Dim vSplit() As Variant
    Dim Branch As String
    Dim ReportDate As String
    Dim BrNumber As String
    Dim Test As Variant
    Dim TotalRows As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Open inventory count sheet
    Path = Application.GetOpenFilename("All Files (*.*), *.*", Title:="Open Inventory Count Sheet (stdin.txt)")
    If Path = "False" Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Exit Sub
    Else
        Workbooks.Open Path
    End If

    'Check to first cell to see if it contains a special character
    'that identifies the sheet as the inventory file
    If Range("A1").Text <> " " Then
        MsgBox "File validation failed." & vbCrLf & "Please make sure you selected the correct inventory file."
        ActiveWorkbook.Close
        Exit Sub
    End If

    'Import inventory count sheet
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Range("A1:A" & TotalRows).Copy Destination:=ThisWorkbook.Sheets("Raw").Range("A1")
    ActiveWorkbook.Close

    'Process the report header to get the date and branch
    Sheets("Temp").Select
    Range("A1").Value = Sheets("Raw").Range("A5").Text
    Range("A1").TextToColumns Destination:=Range("A1"), _
                              DataType:=xlFixedWidth, _
                              FieldInfo:=Array(Array(0, 1), Array(4, 1), Array(51, 1), Array(122, 1), Array(130, 1)), _
                              TrailingMinusNumbers:=True
    ReportDate = Range("B1").Text
    Branch = Range("C1").Text
    ActiveSheet.UsedRange.Delete

    BrNumber = InputBox("Enter your branch number.", "Enter Branch #")
    If BrNumber = "" Then
        BrNumber = "0000"
    End If

    Sheets("Raw").Select
    vArray = ActiveSheet.UsedRange.Rows

    'Clear data from the array that is not needed
    For i = 1 To UBound(vArray)
        If InStr(CStr(vArray(i, 1)), "PHYSICAL INVENTORY") Then
            vArray(i, 1) = ""
        ElseIf InStr(CStr(vArray(i, 1)), "PAGE") Then
            vArray(i, 1) = ""
        ElseIf InStr(CStr(vArray(i, 1)), "CHECKED BY") Then
            vArray(i, 1) = ""
        ElseIf InStr(CStr(vArray(i, 1)), "SIM NUMBER") Then
            vArray(i, 1) = ""
        ElseIf InStr(CStr(vArray(i, 1)), "ITEM DESCRIPTION") Then
            vArray(i, 1) = ""
        ElseIf InStr(CStr(vArray(i, 1)), "COUNTED BY") Then
            vArray(i, 1) = ""
        ElseIf InStr(CStr(vArray(i, 1)), Branch) Then
            vArray(i, 1) = ""
        ElseIf InStr(CStr(vArray(i, 1)), "END OF REPORT") Then
            vArray(i, 1) = ""
        ElseIf InStr(CStr(vArray(i, 1)), "") Then
            vArray(i, 1) = ""
        ElseIf vArray(i, 1) = " " Then
            vArray(i, 1) = ""
        End If
    Next

    'Fill the count sheet with inventory data
    Sheets("Count Sheet").Select
    Range("B1:B" & UBound(vArray)) = vArray

    'Fill column A so that it can be filtered
    Range("A1").Value = "Col1"
    Range("A2:A" & UBound(vArray)).Value = "x"

    'Add a header to column B so it can be filtered
    Range("B1").Value = "Col2"

    'Remove all blank cells
    Range("A1:B" & UBound(vArray)).AutoFilter Field:=2, Criteria1:="<>"
    ActiveSheet.UsedRange.Copy Destination:=Sheets("Temp").Range("A1")
    Sheets("Temp").Columns(1).Delete
    Sheets("Temp").Rows(1).Delete
    vArray = Worksheets("Temp").UsedRange
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Cells.Delete
    Sheets("Temp").Cells.Delete
    Range(Cells(1, 1), Cells(UBound(vArray), 1)) = vArray

    'vArray should always be an even number because
    'inventory lines are added 2 at a time and all
    'other data has been removed
    ReDim vSplit(1 To UBound(vArray) / 2, 1 To 2)
    For i = 1 To UBound(vArray)
        If i Mod 2 = 0 Then
            'Put all even lines in column 1
            vSplit(i / 2, 2) = vArray(i, 1)
        Else
            'put all odd lines in column 2
            x = (i / 2) + 0.5
            vSplit(x, 1) = vArray(i, 1)
        End If
    Next

    Sheets("Count Sheet").Cells.Delete
    Range("A1:B" & UBound(vSplit)) = vSplit

    Columns("B:B").TextToColumns Destination:=Range("B1"), _
                                 DataType:=xlFixedWidth, _
                                 FieldInfo:=Array(Array(0, 1), Array(21, 1), Array(81, 1), Array(90, 1), Array(99, 1), Array(108, 1), Array(117, 1), Array(127, 1)), _
                                 TrailingMinusNumbers:=True

    Range("B:F").EntireColumn.Insert
    Columns("A:A").TextToColumns Destination:=Range("A1"), _
                                 DataType:=xlFixedWidth, _
                                 FieldInfo:=Array(Array(0, 1), Array(2, 1), Array(17, 1), Array(20, 1), Array(28, 1), Array(39, 1)), _
                                 TrailingMinusNumbers:=True

    Rows(1).Insert Shift:=xlDown
    Range("A1:N1").Value = Array("LN #", "SIM NUMBER", "UOM", "CON", "WIP", "WIT", "LOCATION", "ITEM DESCRIPTION", "COUNT   #1", "COUNT TOTAL", "RECHECK  #1", "RECHECK  #2", "DELETE", "DELETE")
    'Reorder columns, add page numbers, align text, and set font
    ReStructure

    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$1"
        .PrintTitleColumns = ""
    End With

    ActiveSheet.PageSetup.PrintArea = ""

    With ActiveSheet.PageSetup
        .LeftHeader = "&15&B " & BrNumber & "  " & Branch & "  &B"
        .CenterHeader = "&15&B" & ReportDate & " Physcial Inventory    &B"
        .RightHeader = _
        "&15&BCounted By:&B _____________________________ " & Chr(10) & "" & Chr(10) & "&BRechecked By:&B ___________________________ "
        .LeftFooter = ""
        .CenterFooter = "&15Page &P of &N"
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(1)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = True
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesTall = False
        .FitToPagesWide = 1
        .PrintErrors = xlPrintErrorsDisplayed
    End With

    Application.DisplayAlerts = True
    Sheets("Count Sheet").Copy
    On Error GoTo SAVE_FAILED
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & BrNumber & " Count Sheet " & Format(Date, "mm-dd-yy"), FileFormat:=xlNormal
    On Error GoTo 0
    MsgBox "Saved to: " & vbCrLf & ThisWorkbook.Path & "\" & BrNumber & " Count Sheet " & Format(Date, "mm-dd-yy") & ".xls"
    CleanUp
    Application.ScreenUpdating = True
    Application.ScreenUpdating = True
    Exit Sub

SAVE_FAILED:
    MsgBox "A copy of the count sheet was NOT saved."
    Application.DisplayAlerts = False
    CleanUp
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

End Sub

Sub ReStructure()
    Dim iPage As Integer
    Dim vArray As Variant
    Dim vSplit() As Variant
    Dim i As Long

    Columns(1).EntireColumn.Insert Shift:=xlToRight
    vArray = Range(Cells(1, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 2))

    For i = 1 To UBound(vArray)
        If vArray(i, 2) = 1 Then
            iPage = iPage + 1
        End If
        vArray(i, 1) = iPage
    Next

    ReDim vSplit(1 To UBound(vArray), 1 To 1)
    For i = 1 To UBound(vArray)
        vSplit(i, 1) = vArray(i, 1)
    Next

    Range(Cells(1, 1), Cells(UBound(vSplit), 1)) = vSplit
    Range("A1").Value = "PG #"
    Range("N:O").Delete Shift:=xlToLeft
    Range(Cells(2, 10), Cells(ActiveSheet.UsedRange.Rows.Count, 13)).Delete Shift:=xlUp

    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 15
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
    End With

    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    Columns("I:I").Cut
    Columns("E:E").Insert Shift:=xlToRight
    Columns("I:I").Cut
    Columns("F:F").Insert Shift:=xlToRight
    Range(Cells(2, 5), Cells(ActiveSheet.UsedRange.Rows.Count, 5)).HorizontalAlignment = xlLeft
    ActiveSheet.UsedRange.RowHeight = 40
    ActiveSheet.UsedRange.EntireColumn.AutoFit
End Sub

Sub CleanUp()
    Dim s As Worksheet
    Dim PrevDispAlert As Boolean
    Dim PrevScrnUpdat As Boolean

    PrevDispAlert = Application.DisplayAlerts
    PrevScrnUpdat = Application.ScreenUpdating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ThisWorkbook.Activate

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.Cells.Delete
        End If
    Next

    Application.DisplayAlerts = PrevDispAlert
    Application.ScreenUpdating = PrevScrnUpdat

    Sheets("Macro").Select
    Range("C7").Select
End Sub
