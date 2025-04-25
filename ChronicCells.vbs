'********************
'* Helper Functions *
'********************
'Function to return the number of used rows on a sheet.
Private Function iRowCount(ThisSheet As Worksheet, iColumn as Integer) As Integer
    Dim lCount As Long
    With ThisSheet
        For lCount = 1 To 1500
            If .Cells(lCount, iColumn).Value <> "" Then
                iRowCount = lCount
            End If
        Next
    End With
End Function
'Function to return the number of used columns on a sheet.
Private Function iColumnCount(ThisSheet As Worksheet, iRow as Integer) As Integer
    Dim lCount As Long
    With ThisSheet
        For lCount = 1 To 1000
            If .Cells(iRow, lCount).Value <> "" Then
                iColumnCount = lCount
            End If
        Next
    End With
End Function
'***************************************
'* Availability Chronic Cells - 7 Days *
'***************************************
Sub AvailabilityChronicCells_7Days()
    Dim wsFirstSheet As Worksheet
    Set wsFirstSheet = Application.ThisWorkbook.Sheets(1)
    Dim iLastRow As Integer
    Rows("1:4").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Rows(1).Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -10921639
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    ActiveCell.FormulaR1C1 = "SITE_ID"
    Rows("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Tickets"
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "Alarms"
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Comments"
    Range("N2").Select
    Columns("N:N").ColumnWidth = 8.57
    Columns("N:N").EntireColumn.AutoFit
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Last 3 Days"
    Range("O2").Select
    Columns("O:O").EntireColumn.AutoFit
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=HYPERLINK(""http://natweb.eng.t-mobile.com/sites/Reporting/Reports/Homer/TTWOINCHistoryReport.aspx?SiteID=""&RC[-11],""Tickets"")"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = _
        "=HYPERLINK(""https://axiom.network.t-mobile.com/site/""&RC[-12],""AXIOM"")"
    Range("M3").Select
    Columns("M:M").EntireColumn.AutoFit
    Range("O2").Select
    ActiveCell.Formula2R1C1 = _
        "=IF(AND(RC[-10]=100%,F=100%,RC[-8]=100%),""GOOD"",""BAD"")"
    Range("O2").Select
    ActiveCell.Formula2R1C1 = _
        "=IF(AND(RC[-10]=100,F=100,RC[-8]=100),""GOOD"",""BAD"")"
    Range("O2").Select
    ActiveCell.Formula2R1C1 = _
        "=IF(AND(RC[-10]=100,F=100,RC[-8]=100),""GOOD"",""BAD"")"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "IF(AND("
    Range("E2").Select
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "IF(AND(E2=100%, F2=100%, G2=100%),""GOOD"",""BAD"")"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "IF(AND(E2=100%,F2=100%,G2=100%),""GOOD"",""BAD"")"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-10]=100%,RC[-9]=100%,RC[-8]=100%),""GOOD"",""BAD"")"
    Range("L2").Select
    Selection.AutoFill Destination:=Range("L2:L11999"), Type:=xlFillDefault
    Range("L2:L1999").Select

    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M1999"), Type:=xlFillDefault
    Range("M2:M1999").Select

    Range("O2").Select
    Selection.AutoFill Destination:=Range("O2:O11999"), Type:=xlFillDefault
    Range("O2:O1999").Select

    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$O$11999").AutoFilter Field:=15, Criteria1:="BAD"

    Sheets("Sheet01").Select
    Sheets("Sheet01").Name = "Availability 7 Days"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "QTD & DAILY DATA"
End Sub
'*********************************
'* Change dashes in data to 100% *
'*********************************
Sub DashTo100()
    Sheets("Availability 7 Days").Select
    Columns("E:K").Select
    Selection.Replace What:="-", Replacement:="100%", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveSheet.Range("$A$1:$O$279").AutoFilter Field:=15, Criteria1:="BAD"
    Columns("A:D").Select
    Columns("A:D").EntireColumn.AutoFit
    ActiveWindow.LargeScroll ToRight:=-1
    Columns("O:O").Select
    Selection.EntireColumn.Hidden = True
    Columns("N:N").Select
    Selection.ColumnWidth = 103
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft

    ActiveSheet.Range("$A$1:$O$11999").AutoFilter Field:=5, Criteria1:="<>"
    ActiveSheet.Range("$A$1:$O$11999").AutoFilter Field:=5, Criteria1:=RGB(128 _
        , 0, 0), Operator:=xlFilterCellColor
End Sub
