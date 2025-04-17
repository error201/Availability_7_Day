Sub AvailabilityChronicCells_7Days()
'
' AvailabilityChronicCells_7Days Macro
' Availability Chronic Cells - 7 Days
'

'
    Rows("1:4").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1:K1").Select
    With Selection.Font
        .Name = "Calibri"
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
    ActiveWindow.SmallScroll Down:=-15
    ActiveWindow.ScrollRow = 149
    ActiveWindow.ScrollRow = 147
    ActiveWindow.ScrollRow = 145
    ActiveWindow.ScrollRow = 144
    ActiveWindow.ScrollRow = 142
    ActiveWindow.ScrollRow = 141
    ActiveWindow.ScrollRow = 139
    ActiveWindow.ScrollRow = 138
    ActiveWindow.ScrollRow = 137
    ActiveWindow.ScrollRow = 135
    ActiveWindow.ScrollRow = 134
    ActiveWindow.ScrollRow = 133
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 129
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 127
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 125
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 123
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 119
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 123
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 119
    ActiveWindow.ScrollRow = 118
    ActiveWindow.ScrollRow = 117
    ActiveWindow.ScrollRow = 116
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 113
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 109
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 105
    ActiveWindow.ScrollRow = 103
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 88
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 80
    ActiveWindow.ScrollRow = 77
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 2
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M1999"), Type:=xlFillDefault
    Range("M2:M1999").Select
    ActiveWindow.SmallScroll Down:=-9
    ActiveWindow.ScrollRow = 157
    ActiveWindow.ScrollRow = 156
    ActiveWindow.ScrollRow = 155
    ActiveWindow.ScrollRow = 153
    ActiveWindow.ScrollRow = 149
    ActiveWindow.ScrollRow = 145
    ActiveWindow.ScrollRow = 140
    ActiveWindow.ScrollRow = 135
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 109
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 94
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    Range("O2").Select
    Selection.AutoFill Destination:=Range("O2:O11999"), Type:=xlFillDefault
    Range("O2:O1999").Select
    ActiveWindow.SmallScroll Down:=-30
    ActiveWindow.ScrollRow = 136
    ActiveWindow.ScrollRow = 135
    ActiveWindow.ScrollRow = 133
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 119
    ActiveWindow.ScrollRow = 117
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 113
    ActiveWindow.ScrollRow = 111
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 105
    ActiveWindow.ScrollRow = 104
    ActiveWindow.ScrollRow = 102
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 66
    ActiveWindow.ScrollRow = 64
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$O$11999").AutoFilter Field:=15, Criteria1:="BAD"
    ActiveWindow.SmallScroll Down:=-27
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 88
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 92
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 94
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 100
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 102
    ActiveWindow.ScrollRow = 103
    ActiveWindow.ScrollRow = 104
    ActiveWindow.ScrollRow = 105
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 109
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 111
    ActiveWindow.ScrollRow = 113
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 116
    ActiveWindow.ScrollRow = 117
    ActiveWindow.ScrollRow = 118
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 122
    ActiveWindow.ScrollRow = 123
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 125
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 127
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 129
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 131
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 133
    ActiveWindow.ScrollRow = 134
    ActiveWindow.ScrollRow = 135
    ActiveWindow.ScrollRow = 136
    ActiveWindow.ScrollRow = 137
    ActiveWindow.ScrollRow = 138
    ActiveWindow.ScrollRow = 139
    ActiveWindow.ScrollRow = 140
    ActiveWindow.ScrollRow = 141
    ActiveWindow.ScrollRow = 142
    ActiveWindow.ScrollRow = 143
    ActiveWindow.ScrollRow = 144
    ActiveWindow.ScrollRow = 145
    ActiveWindow.ScrollRow = 146
    ActiveWindow.ScrollRow = 147
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 149
    ActiveWindow.ScrollRow = 150
    ActiveWindow.ScrollRow = 151
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 153
    ActiveWindow.ScrollRow = 154
    ActiveWindow.ScrollRow = 155
    ActiveWindow.ScrollRow = 156
    ActiveWindow.ScrollRow = 157
    ActiveWindow.ScrollRow = 158
    ActiveWindow.ScrollRow = 159
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 161
    ActiveWindow.ScrollRow = 162
    ActiveWindow.ScrollRow = 163
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 165
    ActiveWindow.ScrollRow = 166
    Sheets("Sheet01").Select
    Sheets("Sheet01").Name = "Availability 7 Days"
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "QTD & DAILY DATA"
End Sub


Sub DashTo100()
'
' DashTo100 Macro
'

'
	Sheets("Availability 7 Days").Select
    Columns("E:K").Select
    Selection.Replace What:="-", Replacement:="100%", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Font.Size = 10
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