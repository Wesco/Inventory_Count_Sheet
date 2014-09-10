Attribute VB_Name = "Program"
Option Explicit

Sub CountSheet()
    Dim i As Long
    Dim x As Long
    Dim sPath As String
    Dim vArray As Variant
    Dim vSplit() As Variant
    Dim sBranch As String
    Dim sReportDate As String
    Dim sBrNumber As String
    Dim sTest As Variant

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    sPath = Application.GetOpenFilename("stdin (*.txt), *.txt")

    If sPath = "False" Then
        Exit Sub
    Else
        Workbooks.Open sPath
    End If

    If Range("A1").Text <> " " Then
        MsgBox "File validation failed." & vbCrLf & "Please make sure you selected the correct inventory file."
        ActiveWorkbook.Close
        Exit Sub
    End If

    Range(Cells(1, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 1)).Copy Destination:=ThisWorkbook.Worksheets("Raw").Range("A1")

    ActiveWorkbook.Close
    Worksheets("Raw").Select

    Worksheets("Temp").Range("A1").Value = Cells(5, 1).Text
    Worksheets("Temp").Select
    Range("A1").Select
    Selection.TextToColumns Destination:=Range("A1"), _
                            DataType:=xlFixedWidth, _
                            FieldInfo:=Array(Array(0, 1), Array(4, 1), Array(51, 1), Array(122, 1), Array(130, 1)), _
                            TrailingMinusNumbers:=True

    sReportDate = Range("B1").Text
    sBranch = Range("C1").Text
    ActiveSheet.UsedRange.Delete


    sBrNumber = InputBox("Enter your branch number.", "Enter Branch #")

    If sBrNumber = "" Then
        sBrNumber = "0000"
    End If

    Worksheets("Raw").Select
    vArray = ActiveSheet.UsedRange.Rows

    For i = 1 To UBound(vArray)
        If InStr(CStr(vArray(i, 1)), "PHYSICAL INVENTORY") Then
            vArray(i, 1) = ""
        End If
        If InStr(CStr(vArray(i, 1)), "PAGE") Then
            vArray(i, 1) = ""
        End If
        If InStr(CStr(vArray(i, 1)), "CHECKED BY") Then
            vArray(i, 1) = ""
        End If
        If InStr(CStr(vArray(i, 1)), "SIM NUMBER") Then
            vArray(i, 1) = ""
        End If
        If InStr(CStr(vArray(i, 1)), "ITEM DESCRIPTION") Then
            vArray(i, 1) = ""
        End If
        If InStr(CStr(vArray(i, 1)), "COUNTED BY") Then
            vArray(i, 1) = ""
        End If
        If InStr(CStr(vArray(i, 1)), sBranch) Then
            vArray(i, 1) = ""
        End If
        If InStr(CStr(vArray(i, 1)), "END OF REPORT") Then
            vArray(i, 1) = ""
        End If
        If InStr(CStr(vArray(i, 1)), "") Then
            vArray(i, 1) = ""
        End If
        If vArray(i, 1) = " " Then
            vArray(i, 1) = ""
        End If
    Next

    Worksheets("Count Sheet").Select
    Range(Cells(1, 2), Cells(UBound(vArray), 2)) = vArray
    Range("A1").Value = "Col1"
    Range(Cells(2, 1), Cells(UBound(vArray), 1)).Value = "x"
    Range("B1").Value = "Col2"
    Range(Cells(1, 1), Cells(UBound(vArray), 2)).AutoFilter Field:=2, Criteria1:="<>"

    ActiveSheet.UsedRange.Copy Destination:=Worksheets("Temp").Range("A1")
    Worksheets("Temp").Columns(1).Delete
    Worksheets("Temp").Rows(1).Delete
    vArray = Worksheets("Temp").UsedRange
    ActiveSheet.AutoFilterMode = False
    ActiveSheet.Cells.Delete
    Worksheets("Temp").Cells.Delete
    Range(Cells(1, 1), Cells(UBound(vArray), 1)) = vArray

    ReDim vSplit(1 To UBound(vArray) / 2, 1 To 2)

    x = 0
    For i = 1 To UBound(vArray)
        If i Mod 2 = 0 Then
            vSplit(i / 2, 2) = vArray(i, 1)
        Else
            x = (i / 2) + 0.5
            vSplit(x, 1) = vArray(i, 1)
        End If
    Next

    Worksheets("Count Sheet").Cells.Delete
    Range(Cells(1, 1), Cells(UBound(vSplit), 2)) = vSplit

    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlFixedWidth, _
                            FieldInfo:=Array(Array(0, 1), Array(21, 1), Array(81, 1), Array(90, 1), Array(99, 1), _
                                             Array(108, 1), Array(117, 1), Array(127, 1)), TrailingMinusNumbers:=True

    Range("B1").Select
    Range("B:F").EntireColumn.Insert
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), _
                            DataType:=xlFixedWidth, _
                            FieldInfo:=Array(Array(0, 1), Array(2, 1), Array(17, 1), Array(20, 1), Array(28, 1), Array(39, 1)), _
                            TrailingMinusNumbers:=True

    Range("A1").Select
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
        .LeftHeader = "&15&B " & sBrNumber & "  " & sBranch & "  &B"
        .CenterHeader = "&15&B" & sReportDate & " Physcial Inventory    &B"
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
    Worksheets("Count Sheet").Copy
    On Error GoTo SAVE_FAILED
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & sBrNumber & " Count Sheet " & Format(Date, "mm-dd-yy"), FileFormat:=xlNormal
    On Error GoTo 0
    MsgBox "Saved to: " & vbCrLf & ThisWorkbook.Path & "\" & sBrNumber & " Count Sheet " & Format(Date, "mm-dd-yy") & ".xls"
    Application.DisplayAlerts = False
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
    ThisWorkbook.Worksheets("Raw").Cells.Delete
    ThisWorkbook.Worksheets("Temp").Cells.Delete
    ThisWorkbook.Worksheets("Count Sheet").Cells.Delete
    ThisWorkbook.Save
End Sub






