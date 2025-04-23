'Function to return the number of used rows on a sheet.
Private Function iRowCount(ThisSheet As Worksheet) As Integer
    Dim lCount As Long
    With ThisSheet
        For lCount = 1 To 1500
            If .Cells(lCount, 1).Value <> "" Then
                iRowCount = lCount
            End If
        Next
    End With
End Function
'Function to return the number of used columns on a sheet.
Private Function iColumnCount(ThisSheet As Worksheet) As Integer
    Dim lCount As Long
    With ThisSheet
        For lCount = 1 To 1000
            If .Cells(1, lCount).Value <> "" Then
                iColumnCount = lCount
            End If
        Next
    End With
End Function
'Function to determine if a value exists in an array (set membership)
Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function

Private Function boolGreaterThanOneDay(strString As Variant) As Boolean
    Dim regexOne As Object
    Dim strMyString As String
    Set regexOne = New RegExp
    regexOne.Global = False
    regexOne.Pattern = "for (\d{1,3}) D"
    Set theMatches = regexOne.Execute(strString)
    For Each Match In theMatches
        If Match.SubMatches.Count > 0 Then
            For Each subMatch In Match.SubMatches
                If subMatch <> "0" Then
                    boolGreaterThanOneDay = True
                    Exit Function
                End If
            Next subMatch
        End If
    Next
    boolGreaterThanOneDay = False
End Function

Sub newOne()
    'This subroutine is assigned to the hotkey CRTL+Shift+W
    Set OneC = Sheets("3 - 1C Export")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Set common colors
    rgbMagenta = RGB(226, 0, 116)
    rgbPaleYellow = RGB(255, 230, 153)
    rgbMedGray = RGB(128, 128, 128)
    rgbStatusUp = RGB(146, 208, 80)
    rgbStatusDown = RGB(232, 48, 61)
    rgbStatusSector = RGB(246, 117, 22)
    rgbStatusOther = RGB(0, 0, 0)
    rgbStatusPriority = RGB(0, 176, 240)
    
    'Subroutine variables
    Dim cl As Long
    Dim thiscl As Long
    
    Dim iColumnCounter
    Dim iRowCounter
    Dim strValue
    Dim strThisValue
    Dim iMaxRow
    Dim iMaxColumn
    'Dim rngCurrentCell As Range
    'Dim rngCurrentRow As Range
    'Dim rngCurrentColumn As Range
    Dim arrAllowedColumns As Variant
    arrAllowedColumns = Array("Config Item", "CI Name", "GSM", "LTE", "5G", "Ticket ID", "TT Count", "TT Status", "Cause", "Default Tech", "AAVVendor")
    iMaxRow = iRowCount(Application.Worksheets("3 - 1C Export"))
    iMaxColumn = iColumnCount(Application.Worksheets("3 - 1C Export"))
    
OneC.Activate
    If OneC.Cells(1, 1).Value Like "Paste*" Then
        MsgBox ("1C Export sheet is empty. Paste export from OneConsole to continue.")
        Exit Sub
    End If
    If Cells(1, 1).Value = "Site ID/ Type" Then Rows(1).EntireRow.Delete
    If Cells(1, 1).Value = "Sites Down <24 Hours" Then Exit Sub


    'Re-order Columns
    Dim ColumnOrder As Variant, ndx As Integer
    Dim Found As Range, counter As Integer
        ColumnOrder = Array("Config Item", "CI Name", "GSM", "LTE", "5G", "Ticket ID", "TT Count", "TT Status", "Cause", "Default Tech", "AAVVendor")
    counter = 1
    For ndx = LBound(ColumnOrder) To UBound(ColumnOrder)
        Set Found = Rows("1:1").Find(ColumnOrder(ndx), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not Found Is Nothing Then
            If Found.Column <> counter Then
                Found.EntireColumn.Cut
                Columns(counter).Insert Shift:=xlToRight
                Application.CutCopyMode = False
            End If
        counter = counter + 1
        End If
    Next
    
    'Find any sites in maintenance status and color their rows
    For iColumnCounter = 1 To iMaxColumn Step 1
        strValue = Cells(iColumnCounter).Value
        If strValue = "" Then Exit For
        If strValue = "GSM NEST Status" Or strValue = "LTE NEST Status" Or strValue = "5G NEST Status" Then
            For iRowCounter = 1 To iMaxRow Step 1
                strThisValue = Cells(iRowCounter, iColumnCounter)
                If strThisValue = "" Then Exit For
                If InStr(strThisValue, "Maint") > 0 Then
                    'Cells(iRowCounter, iColumnCounter).EntireRow.Interior.Color = rgbPaleYellow
                    Range("A" & iRowCounter & ":M" & iRowCounter).Interior.Color = rgbPaleYellow
                End If
            Next
        End If
    Next
    
    
    'Iterate through the column names...
    For iColumnCounter = 1 To iMaxColumn Step 1
        strValue = Cells(iColumnCounter).Value
        If strValue = "" Then Exit For
        
        'Delete any non-allowed columns.
        If Not IsInArray(strValue, arrAllowedColumns) Then
            Columns(iColumnCounter).EntireColumn.Delete
            iColumnCounter = iColumnCounter - 1
            
        'Format the allowed columns.
        ElseIf strValue = "TT Count" Then
            Cells(iColumnCounter).Value = "#TTs"
            Cells(iColumnCounter).EntireColumn.HorizontalAlignment = xlHAlignCenter
            Cells(iColumnCounter).EntireColumn.Font.Color = rgbBlack
            For iRowCounter = 2 To iMaxRow Step 1
                If Cells(iRowCounter, iColumnCounter).Interior.Color = rgbPaleYellow Then
                    Cells(iRowCounter, iColumnCounter).Interior.Color = rgbPaleYellow
                Else
                    Cells(iRowCounter, iColumnCounter).Interior.Color = rgbWhite
                End If
            Next
            
            
        ElseIf strValue = "AAVVendor" Then
            Cells(iColumnCounter).Value = "AAV"
            For iRowCounter = 1 To iMaxRow Step 1
                strThisValue = Cells(iRowCounter, iColumnCounter)
                If strThisValue = "" Then Exit For
                
                If InStr(strThisValue, "Not Available") > 0 Then
                    strThisValue = "NA"
                End If
                
                Cells(iRowCounter, iColumnCounter).Value = strThisValue
            Next

        ElseIf strValue = "GSM" Or strValue = "LTE" Or strValue = "5G" Then
            Cells(iColumnCounter).EntireColumn.HorizontalAlignment = xlHAlignCenter
            For iRowCounter = 1 To iMaxRow Step 1
                strThisValue = Cells(iRowCounter, iColumnCounter)
                If strThisValue = "" Then Exit For
                
                If InStr(strThisValue, "DOWN") > 0 Then
                    Cells(iRowCounter, iColumnCounter).Interior.Color = rgbStatusDown
                    Cells(iRowCounter, iColumnCounter).Font.Color = RGB(255, 255, 255)
                    If boolGreaterThanOneDay(strThisValue) Then
                        With Cells(iRowCounter, 2)
                            .Interior.Color = rgbDarkRed
                            .Font.Color = RGB(255, 255, 255)
                        End With
                    End If
                                        
                ElseIf InStr(strThisValue, "UP") > 0 Then
                    Cells(iRowCounter, iColumnCounter).Interior.Color = rgbStatusUp
                    Cells(iRowCounter, iColumnCounter).Font.Color = RGB(0, 0, 0)
                    
                ElseIf InStr(strThisValue, "Sector") > 0 Then
                    Cells(iRowCounter, iColumnCounter).Interior.Color = rgbStatusSector
                    If boolGreaterThanOneDay(strThisValue) Then
                        With Cells(iRowCounter, 2)
                            .Interior.Color = rgbDarkRed
                            .Font.Color = RGB(255, 255, 255)
                        End With
                    End If
 
                ElseIf InStr(UCase(strThisValue), "LOCK") > 0 Then
                    Cells(iRowCounter, iColumnCounter).Interior.Color = RGB(91, 155, 213)
                    Cells(iRowCounter, iColumnCounter).Font.Color = RGB(255, 255, 255)
                    If boolGreaterThanOneDay(strThisValue) Then
                        With Cells(iRowCounter, 2)
                            .Interior.Color = rgbDarkRed
                            .Font.Color = RGB(255, 255, 255)
                        End With
                    End If
                    
                ElseIf InStr(strThisValue, "Not Disc") > 0 Then
                    Cells(iRowCounter, iColumnCounter).Interior.Color = RGB(51, 63, 79)
                    Cells(iRowCounter, iColumnCounter).Font.Color = RGB(255, 255, 255)
                    If boolGreaterThanOneDay(strThisValue) Then
                        With Cells(iRowCounter, 2)
                            .Interior.Color = rgbDarkRed
                            .Font.Color = RGB(255, 255, 255)
                        End With
                    End If
                    
                ElseIf InStr(UCase(strThisValue), "BARRED") > 0 Then
                    Cells(iRowCounter, iColumnCounter).Interior.Color = rgbMagenta
                    Cells(iRowCounter, iColumnCounter).Font.Color = RGB(255, 255, 255)
                    If boolGreaterThanOneDay(strThisValue) Then
                        With Cells(iRowCounter, 2)
                            .Interior.Color = rgbDarkRed
                            .Font.Color = RGB(255, 255, 255)
                        End With
                    End If
                    
                ElseIf InStr(strThisValue, "Not Available") > 0 Then
                    strThisValue = "NA"
                    Cells(iRowCounter, iColumnCounter).Interior.Color = rgbMedGray
                End If
                Cells(iRowCounter, iColumnCounter).Value = strThisValue
            Next
            
        ElseIf strValue = "Cause" Then
            For iRowCounter = 1 To iMaxRow Step 1
                strThisValue = Cells(iRowCounter, iColumnCounter)
                If strThisValue = "" Then Exit For
                
                If InStr(strThisValue, "Not Available") > 0 Then
                    strThisValue = "NA"
                End If
                Cells(iRowCounter, iColumnCounter).Value = strThisValue
            Next
        
        End If
            
    Next
    
    
    'Find any priority sites and color their rows.
    sltSiteCol = 6
    lastrow = Worksheets("Info").Cells(Rows.Count, sltSiteCol).End(xlUp).Row
    For cl = lastrow To 2 Step -1
        If Not Worksheets("Info").Cells(cl, sltSiteCol).Value = "" Then
            sltSite = Worksheets("Info").Cells(cl, sltSiteCol).Value
            OneClastrow = OneC.Cells(Rows.Count, 1).End(xlUp).Row
                For thiscl = 2 To OneClastrow Step 1
                    If Not Cells(thiscl, 1).Value = "" And Cells(thiscl, 1).Value = sltSite Then
                        Cells(thiscl, 1).Interior.Color = rgbStatusPriority
                        thiscl = OneClastrow + 1
                    End If
                Next thiscl
        End If
    Next cl
    
        
    Dim X As Integer
    'Add 2 columns at the end named ETR and Comments.
    X = ActiveSheet.UsedRange.Columns.Count
    Y = ActiveSheet.UsedRange.Rows.Count
    ActiveSheet.Cells(1, X + 1) = "ETR"
    ActiveSheet.Cells(1, X + 2) = "Comments"
    
    For iRowCounter = 2 To iMaxRow Step 1
        If Cells(iRowCounter, 11).Interior.Color = rgbPaleYellow Then
            Range("L" & iRowCounter & ":M" & iRowCounter).Interior.Color = rgbPaleYellow
        End If
    Next
    
    

    'Cells format
    With Range("A1:M1")
        .Font.Bold = True
        .Interior.Color = rgbMagenta
        .Font.Color = rgbWhite
    End With
    
    Cells(1, 1).CurrentRegion.Select
    Selection.Borders.LineStyle = xlContinuous
    
    Columns("A:N").Font.Size = 8
    Columns("A:N").HorizontalAlignment = xlHAlignLeft
    
    Columns("A:Z").Font.Name = "Calibri"
    Columns("A:Z").AutoFit
    
    Columns("L").NumberFormat = "mm/dd/yyyy"
    Columns("L").ColumnWidth = 10
    
    Columns("M").ColumnWidth = 50
    
    Columns("O").ColumnWidth = 5
    

'Remove UMTS column from Previous 1C report sheet
    
    If Worksheets("1 - Previous 1C Report").Cells(1, 4).Value = "UMTS" Or Worksheets("1 - Previous 1C Report").Cells(2, 4).Value = "UMTS" _
    Or Worksheets("1 - Previous 1C Report").Cells(3, 4).Value = "UMTS" Then Worksheets("1 - Previous 1C Report").Columns(4).EntireColumn.Delete

'Remove UMTS column from 24 Hrs report report sheet
    
    If Worksheets("2 - 24 Hrs Report").Cells(1, 4).Value = "UMTS" Or Worksheets("2 - 24 Hrs Report").Cells(2, 4).Value = "UMTS" _
    Or Worksheets("2 - 24 Hrs Report").Cells(3, 4).Value = "UMTS" Then Worksheets("1 - Previous 1C Report").Columns(4).EntireColumn.Delete


'Paste VLOOKUP to ETR and Comments columns
    Set Formula = Sheets("Formula")
    OneClastrow = OneC.Cells(Rows.Count, 1).End(xlUp).Row

        If Cells(1, 12).Value = "ETR" Then
            Formula.Range("B3:C3").Copy
            OneC.Range("L" & 2 & ":L" & OneClastrow).PasteSpecial xlPasteFormulas
            OneC.Range("L" & 2 & ":M" & OneClastrow).Copy
            OneC.Range("L" & 2 & ":L" & OneClastrow).PasteSpecial xlPasteValues
        ElseIf Not Cells(1, 12).Value = "ETR" Then
            MsgBox ("ETR column is not in the correct order or not found")
            Exit Sub
        End If
        
        If Not Worksheets("2 - 24 Hrs Report").Cells(3, 2) = "" Then
            lastrow = Worksheets("2 - 24 Hrs Report").Cells(Rows.Count, 1).End(xlUp).Row
            'MsgBox (lastrow)
            For cl = 1 To OneClastrow Step 1
                If OneC.Cells(cl, 2).Interior.Color = RGB(139, 0, 0) Then
                    SiteId = OneC.Cells(cl, 1).Value
                    For thiscl = 2 To lastrow Step 1
                        If Worksheets("2 - 24 Hrs Report").Cells(thiscl, 1).Value = SiteId Then
                            Formula.Range("B10:C10").Copy
                            OneC.Range("L" & cl & ":L" & cl).PasteSpecial xlPasteFormulas
                        End If
                    Next thiscl
                End If
            Next cl
        End If
        
        OneC.Range("L" & 2 & ":M" & OneClastrow).Copy
        OneC.Range("L" & 2 & ":L" & OneClastrow).PasteSpecial xlPasteValues
'End - Paste VLOOKUP to ETR and Comments columns


'Insert - Sites down <24 hours on top of sheet
    Range("A1").EntireRow.Insert
    With OneC.Cells(1, 1)
        .Value = "Seattle"
        .Font.Bold = "True"
        .Font.Size = "11"
    End With
    
    Range("A1").EntireRow.Insert
    With OneC.Cells(1, 1)
        .Value = "Sites Down <24 Hours"
        .Font.Bold = "True"
        .Font.Size = "11"
    End With


'INW Section - Start
    OneClastrow = OneC.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Insert section name
    With OneC.Cells(OneClastrow + 2, 1)
        .Value = "INW"
        .Font.Bold = "True"
        .Font.Size = "11"
    End With
    
    'Check for header then copy and insert
    If Cells(3, 1).Value = "Config Item" Then
        Range("A3:N3").Copy
        Cells(Cells(Rows.Count, 1).End(xlUp).Row + 1, 1).Insert
    End If
    
    For cl = OneClastrow To 1 Step -1
        If Not OneC.Cells(cl, 2).Interior.Color = RGB(139, 0, 0) Then
            If Cells(cl, 1).Value Like "MT*" Or Cells(cl, 1).Value Like "SP*" Then
                Cells(cl, 1).EntireRow.Cut Cells(Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
                Cells(cl, 1).EntireRow.Delete
                cl = cl + 1
            End If
        End If
    Next cl
    
    OneClastrow = OneC.Cells(Rows.Count, 1).End(xlUp).Row
    For cl = OneClastrow To 1 Step -1
        If Cells(cl, 1).Value = "INW" Then
            Range("A" & cl + 1 & ":N" & OneClastrow).Sort Key1:=Range("A" & cl + 1 & ":A" & OneClastrow), Order1:=xlAscending, Header:=xlYes
            cl = 0
        End If
    
    Next cl
       
'INW Section - End


'PT Section - Start
    OneClastrow = OneC.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Insert section name
    With OneC.Cells(OneClastrow + 2, 1)
        .Value = "Portland"
        .Font.Bold = "True"
        .Font.Size = "11"
    End With
    
    'Check for header then copy and insert
    If Cells(3, 1).Value = "Config Item" Then
        Range("A3:N3").Copy
        Cells(Cells(Rows.Count, 1).End(xlUp).Row + 1, 1).Insert
    End If
    
    For cl = OneClastrow To 1 Step -1
        If Not OneC.Cells(cl, 2).Interior.Color = RGB(139, 0, 0) Then
            If Cells(cl, 1).Value Like "PO*" Then
                Cells(cl, 1).EntireRow.Cut Cells(Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
                Cells(cl, 1).EntireRow.Delete
                cl = cl + 1
            End If
        End If
    Next cl
    
    OneClastrow = OneC.Cells(Rows.Count, 1).End(xlUp).Row
    For cl = OneClastrow To 1 Step -1
        If Cells(cl, 1).Value = "Portland" Then
            Range("A" & cl + 1 & ":N" & OneClastrow).Sort Key1:=Range("A" & cl + 1 & ":A" & OneClastrow), Order1:=xlAscending, Header:=xlYes
            cl = 0
        End If
    
    Next cl
       
'PT Section - End


'>24 Hours Section - Start
    OneClastrow = OneC.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Insert section name
    With OneC.Cells(OneClastrow + 2, 1)
        .Value = "Sites Down >24 Hours"
        .Font.Bold = "True"
        .Font.Size = "11"
    End With
    
    'Check for header then copy and insert
    If Cells(3, 1).Value = "Config Item" Then
        Range("A3:N3").Copy
        Cells(Cells(Rows.Count, 1).End(xlUp).Row + 1, 1).Insert
    End If
    
    For cl = OneClastrow To 1 Step -1
        If OneC.Cells(cl, 2).Interior.Color = RGB(139, 0, 0) Then
                Cells(cl, 1).EntireRow.Cut Cells(Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
                Cells(cl, 1).EntireRow.Delete
                cl = cl + 1
        End If
    Next cl
    
    OneClastrow = OneC.Cells(Rows.Count, 1).End(xlUp).Row
    For cl = OneClastrow To 1 Step -1
        If Cells(cl, 1).Value = "Sites Down >24 Hours" Then
            Range("A" & cl + 1 & ":N" & OneClastrow).Sort Key1:=Range("A" & cl + 1 & ":A" & OneClastrow), Order1:=xlAscending, Header:=xlYes
            cl = 0
        End If
    
    Next cl
       
'>24 Hours Section - End


'Add a legend for management
    OneClastrow = OneC.Cells(Rows.Count, 1).End(xlUp).Row
    
    Cells(OneClastrow + 3, 1).Value = "Cell Color"
    Cells(OneClastrow + 3, 2).Value = "Meaning"
    
    Cells(OneClastrow + 4, 1).Interior.Color = rgbDarkRed
    Cells(OneClastrow + 4, 2).Value = "Down > 24 Hours"
    Cells(OneClastrow + 4, 2).Font.Size = "10"
    
    Cells(OneClastrow + 5, 1).Interior.Color = rgbPaleYellow
    Cells(OneClastrow + 5, 2).Value = "NESTED by Operations"
    Cells(OneClastrow + 5, 2).Font.Size = "10"
    
    Cells(OneClastrow + 6, 1).Interior.Color = RGB(244, 176, 132)
    Cells(OneClastrow + 6, 2).Value = "NESTED by Other"
    Cells(OneClastrow + 6, 2).Font.Size = "10"
    
    Cells(OneClastrow + 7, 1).Interior.Color = rgbStatusPriority
    Cells(OneClastrow + 7, 2).Value = "Priority. SLT site."
    Cells(OneClastrow + 7, 2).Font.Size = "10"
    
    Cells(OneClastrow + 3, 1).CurrentRegion.Select
    With Selection
        .Borders.LineStyle = xlContinuous
    End With
    
    With Range("A" & OneClastrow + 3 & ":B" & OneClastrow + 3)
        .Font.Bold = True
        .Interior.Color = rgbMagenta
        .Font.Color = rgbWhite
        .Font.Size = "10"
        .HorizontalAlignment = xlHAlignCenter
    End With
   
   
   'Range("N:AZ").ClearFormats
   

'Add guidelines section
    OneClastcol = OneC.Cells(5, Columns.Count).End(xlToLeft).Column - 2
    
    'ActiveSheet.Cells(row#, X + "columns# from Default Tech") = "text"
    OneC.Cells(1, OneClastcol + 4) = "Comment Guidelines and Structure:"
    OneC.Cells(2, OneClastcol + 4) = "Part 1 - Alarm: "
    OneC.Cells(2, OneClastcol + 5) = "Current alarms and issues"
    OneC.Cells(3, OneClastcol + 4) = "Part 2 - Action: "
    OneC.Cells(3, OneClastcol + 5) = "Action taken by SWOPS, FOPS, NOC, RF, etc… or being taken to resolve the issues"
    OneC.Cells(4, OneClastcol + 4) = "Part 3 - Time: "
    OneC.Cells(4, OneClastcol + 5) = "ETA or ETR if available"
    OneC.Cells(6, OneClastcol + 4) = "Examples: "
    OneC.Cells(6, OneClastcol + 5) = "Alpha AHFIG down. TT created and assigned to FOP. FOPS notified."
    OneC.Cells(7, OneClastcol + 5) = "Beta AHLOA down. TC being scheduled by FOPS. Tentatively expected 01/01."
    OneC.Cells(8, OneClastcol + 5) = "LTE down. Access requires advanced notice. FOPS arranging access. Tentatively expected 01/01."
    OneC.Cells(10, OneClastcol + 4) = "Please refrain from using the same comments"
    OneC.Cells(11, OneClastcol + 4) = "as previous report, EXCEPT for: "
    OneC.Cells(11, OneClastcol + 5) = "ETA has been confirmed and posted on TT."
    OneC.Cells(12, OneClastcol + 5) = "Sites in maintenance / NEST."
    OneC.Cells(13, OneClastcol + 5) = "Sites in the daily 24 hrs report."
    
    colIDstart = Split(OneC.Cells(1, OneClastcol + 4).Address, "$")(1)
    colIDend = Split(OneC.Cells(1, OneClastcol + 5).Address, "$")(1)

    With Range(colIDstart & 1 & ":" & colIDend & 13)
        .Font.Name = "Calibri"
        .Font.Bold = False
        .Interior.Color = RGB(244, 176, 132)
        .Font.Color = rgbBlack
        .HorizontalAlignment = xlHAlignLeft
    End With
    
    With Range(colIDstart & 1 & ":" & colIDstart & 13)
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignRight
        .ColumnWidth = 40
    End With
    
    Range(colIDend & 1 & ":" & colIDend & 13).ColumnWidth = 80
    
'End - guidelines section

    ActiveWindow.DisplayGridlines = False
    Range("M4").Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    
End Sub
