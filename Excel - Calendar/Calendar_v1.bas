Sub setup()
' Version 1.1, 2014 Jan 23
' Setup the standard blank calendar, creating 2 months in one sheet.
'
'
    yr = InputBox("which year? (yyyy)", , Year(Now))    ' default is this year, "yyyy"
    Do
        mon = Month(Now): mon = InputBox("which month? (1-12)", , ((mon Mod 2) + mon + 1) Mod 12)   ' default is the next odd month, "m"
        If IsNumeric(mon) Then
            If mon > 0 And mon < 13 Then Exit Do
        End If
        MsgBox "Invalid input, please try again."
    Loop

    ref = DateSerial(yr, mon, 1)
    first = DatePart("w", ref, vbMonday)     ' assign the first weekday of the month
    sn = DatePart("ww", ref, vbMonday)       ' assign the first week number of the month
    leap = DatePart("d", DateSerial(yr, 3, 1) - 1)   ' calculate the last day of February
        
    title1 = "'" & MonthName(mon) & " " & yr    ' assign headline title and ending date
    ref = DateSerial(yr, mon + 1, 1)
    title2 = "'" & MonthName(DatePart("m", ref)) & " " & DatePart("yyyy", ref)
    
    end0 = DatePart("d", DateSerial(yr, mon, 1) - 1)
    end1 = DatePart("d", DateSerial(yr, mon + 1, 1) - 1)
    end2 = DatePart("d", DateSerial(yr, mon + 2, 1) - 1)
    
'''''''''' following render the date for each month ''''''''''''
    r = 3: c = 2
    For i = (end0 - first + 2) To end0  ' preceding month of first month
        Cells(r, c) = i
        If c = 8 Then r = r + 1: c = 1
        c = c + 1
    Next
    c0 = c - 1  ' end of first gray

    For i = 1 To end1   ' first month in calendar
        Cells(r, c) = i
        If c = 8 Then r = r + 1: c = 1
        c = c + 1
    Next
    If c = 2 Then c1 = 9 Else c1 = c  ' start of second gray
    first = c - 1   ' first day for next month

    i = 1   ' succeeding month of first month
    Do Until c = 2
        Cells(r, c) = i
        If c = 8 Then r = r + 1: c = 1
        c = c + 1: i = i + 1
    Loop

    line1 = r   'the line number of head2
    r = r + 1

    For i = (end1 - first + 2) To end1  ' preceding month of second month
        Cells(r, c) = i
        If c = 8 Then r = r + 1: c = 1
        c = c + 1
    Next
    c2 = c - 1  ' end of 3rd gray

    For i = 1 To end2   ' second month in calendar
        Cells(r, c) = i
        If c = 8 Then r = r + 1: c = 1
        c = c + 1
    Next
    If c = 2 Then c3 = 9 Else c3 = c ' start of 4th gray

    i = 1   ' succeeding month of second month
    Do Until c = 2
        Cells(r, c) = i
        If c = 8 Then r = r + 1: c = 1
        c = c + 1: i = i + 1
    Loop
    line2 = r   ' the line number of title2
    
''''''''''' following give the general information for layout '''''
    head1 = "B1:H1": head2 = "B" & line1 & ":H" & line1
    body1 = "B3:H" & (line1 - 1): body2 = "B" & (line1 + 1) & ":H" & (line2 - 1)
    days1 = "B2:H2": days2 = "B" & line2 & ":H" & line2
    frame = "B1:H" & line2
    weekend = "G1:H" & line2
    weeksn = "A1:A" & line2
    wks = line2 - 4
        
    Range("a2").ColumnWidth = 2.14
    Range("b2:f2").ColumnWidth = 18.14
    Range("g2:h2").ColumnWidth = 6.71
    
    Range(body1).RowHeight = Round(790 / wks)
    Range(body2).RowHeight = Round(790 / wks)
    
    Range(frame).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    
    For i = 7 To 12         ' border consts, render borders of calendar body
        With Selection.Borders(i)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.25
            .Weight = xlThin
        End With
    Next
        
'''''''' following render the weekend color ''''''''''
    With Range(weekend).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.6
        .PatternTintAndShade = 0
    End With
    
'''''''' Following change the gray color for beginning and ending '''''
    For i = 2 To c0
        Cells(3, i).Interior.ThemeColor = xlThemeColorDark1
        Cells(3, i).Interior.TintAndShade = -0.15
        Cells(3, i).Font.Color = -11711155
    Next
    For i = c1 To 8
        Cells(line1 - 1, i).Interior.ThemeColor = xlThemeColorDark1
        Cells(line1 - 1, i).Interior.TintAndShade = -0.15
        Cells(line1 - 1, i).Font.Color = -11711155
    Next
    For i = 2 To c2
        Cells(line1 + 1, i).Interior.ThemeColor = xlThemeColorDark1
        Cells(line1 + 1, i).Interior.TintAndShade = -0.15
        Cells(line1 + 1, i).Font.Color = -11711155
    Next
    For i = c3 To 8
        Cells(line2 - 1, i).Interior.ThemeColor = xlThemeColorDark1
        Cells(line2 - 1, i).Interior.TintAndShade = -0.15
        Cells(line2 - 1, i).Font.Color = -11711155
    Next


''''''''' Following change the heading year / month information '''''''
    Range(head1).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = True
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.8
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Name = "Consolas"
        .Size = 9
        .Bold = True
    End With
    Range(head1).Cells(1) = title1
    
    Range(head2).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = True
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.8
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Name = "Consolas"
        .Size = 9
        .Bold = True
    End With
    Range(head2).Cells(1) = title2
    
    With Range(days1)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        
        For i = 1 To 5
            .Cells(i) = WeekdayName(i, , vbMonday)
        Next
        For i = 6 To 7
            .Cells(i) = UCase(WeekdayName(i, True, vbMonday))
        Next
    End With
    
    With Range(days2)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        
        For i = 1 To 5
            .Cells(i) = WeekdayName(i, , vbMonday)
        Next
        For i = 6 To 7
            .Cells(i) = UCase(WeekdayName(i, True, vbMonday))
        Next
    End With
    
    Range(weeksn).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
    End With
    With Selection.Font
        .Size = 8
        .ThemeColor = xlThemeColorLight2
    End With

''''''''' following fill the week sequence number ''''''
    For i = 3 To (line1 - 1)
        Cells(i, 1) = sn + 0
        sn = sn + 1
    Next
    
    If Cells(line1 - 1, 8) < 10 Then sn = sn - 1
    If mon = 12 Then sn = 1     ' week sn for next year should start from 1
    
    For i = line1 + 1 To (line2 - 1)
        Cells(i, 1) = sn
        sn = sn + 1
    Next
        
''''''''' following adjust the printing layout ''''''''''
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.17)
        .RightMargin = Application.InchesToPoints(0.17)
        .TopMargin = Application.InchesToPoints(0.17)
        .BottomMargin = Application.InchesToPoints(0.17)
        .Orientation = xlPortrait
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    Application.PrintCommunication = True

End Sub


Sub Appointment()
' Version 1.0, 2014 Jan 22
'
' To make appointment, support 5 types: 1,weekly repeated events, GRAY notes; 2,BLACK notes; 3,BLUE notes;
'   4,RED notes; 5,holiday, RED notes + ORANGE background.
'
' When append new notes, the existed content will be kept, with the assigned note color.
'
'
    
    thm = InputBox("[1] - REPEAT" & vbCrLf & "[2] - Black" & vbCrLf & _
        "[3] - Blue" & vbCrLf & "[4] - Red" & vbCrLf & "[5] - Holiday")
    note = InputBox("comments?", , "_")
    
    Select Case thm
        Case 1
            frq = InputBox("frequent? e.g., [2] for every 2 weeks", , 2)
            With ActiveCell
                arr = Split(.FormulaR1C1, Chr(10))  ' arr is the array keep the text
                ub = UBound(arr): st = 1
                ReDim arr1(ub)  ' arr1 is the arry keep the font color
                For j = 0 To ub
                    l = Len(arr(j))
                    arr1(j) = .Characters(Start:=st, Length:=1).Font.Color
                    st = st + l + 1
                Next
                
                .FormulaR1C1 = .FormulaR1C1 & Chr(10) & note
                
                st = 1
                For j = 0 To ub
                    l = Len(arr(j))
                    .Characters(Start:=st, Length:=l).Font.Color = arr1(j)
                    st = st + l + 1
                Next
                
                .Characters(Start:=st, Length:=Len(note)).Font.Color = -4934476     ' Gray Color
                rw = .Row: rw1 = rw     ' Grab the row number to locate the current week
                wk = Cells(rw, 1)       ' Get the weeksn from column "A"
            End With
            
            For i = rw To 14
                If (Cells(i, 1) = wk + frq) Then
                    ActiveCell.Offset(i - rw1).Activate
                    With ActiveCell
                        If (.Interior.Color = 16777215) Then
                            arr = Split(.FormulaR1C1, Chr(10))  ' arr is the array keep the text
                            ub = UBound(arr): st = 1
                            ReDim arr1(ub)  ' arr1 is the arry keep the font color
                            For j = 0 To ub
                                l = Len(arr(j))
                                arr1(j) = .Characters(Start:=st, Length:=1).Font.Color
                                st = st + l + 1
                            Next
                
                            .FormulaR1C1 = .FormulaR1C1 & Chr(10) & note
                
                            st = 1
                            For j = 0 To ub
                                l = Len(arr(j))
                                .Characters(Start:=st, Length:=l).Font.Color = arr1(j)
                                st = st + l + 1
                            Next
                
                            .Characters(Start:=st, Length:=Len(note)).Font.Color = -4934476     ' Gray Color
                        End If
                        If (.Interior.Color <> 14277081) Then wk = wk + frq
                    End With
                    rw1 = i
                End If
            Next
        Case 2
            With ActiveCell
                arr = Split(.FormulaR1C1, Chr(10))  ' arr is the array keep the text
                ub = UBound(arr): st = 1
                ReDim arr1(ub)  ' arr1 is the arry keep the font color
                For j = 0 To ub
                    l = Len(arr(j))
                    arr1(j) = .Characters(Start:=st, Length:=1).Font.Color
                    st = st + l + 1
                Next
                
                .FormulaR1C1 = .FormulaR1C1 & Chr(10) & note
                
                st = 1
                For j = 0 To ub
                    l = Len(arr(j))
                    .Characters(Start:=st, Length:=l).Font.Color = arr1(j)
                    st = st + l + 1
                Next
                
                .Characters(Start:=st, Length:=Len(note)).Font.Color = -16777216    ' Black color
            End With
        Case 3
            With ActiveCell
                arr = Split(.FormulaR1C1, Chr(10))  ' arr is the array keep the text
                ub = UBound(arr): st = 1
                ReDim arr1(ub)  ' arr1 is the arry keep the font color
                For j = 0 To ub
                    l = Len(arr(j))
                    arr1(j) = .Characters(Start:=st, Length:=1).Font.Color
                    st = st + l + 1
                Next
                
                .FormulaR1C1 = .FormulaR1C1 & Chr(10) & note
                
                st = 1
                For j = 0 To ub
                    l = Len(arr(j))
                    .Characters(Start:=st, Length:=l).Font.Color = arr1(j)
                    st = st + l + 1
                Next
                
                .Characters(Start:=st, Length:=Len(note)).Font.Color = -4165632     ' blue color
            End With
        Case 4
            With ActiveCell
                arr = Split(.FormulaR1C1, Chr(10))  ' arr is the array keep the text
                ub = UBound(arr): st = 1
                ReDim arr1(ub)  ' arr1 is the arry keep the font color
                For j = 0 To ub
                    l = Len(arr(j))
                    arr1(j) = .Characters(Start:=st, Length:=1).Font.Color
                    st = st + l + 1
                Next
                
                .FormulaR1C1 = .FormulaR1C1 & Chr(10) & note
                
                st = 1
                For j = 0 To ub
                    l = Len(arr(j))
                    .Characters(Start:=st, Length:=l).Font.Color = arr1(j)
                    st = st + l + 1
                Next
                
                .Characters(Start:=st, Length:=Len(note)).Font.Color = -16776961    ' red color
            End With
        Case 5
            With ActiveCell
                arr = Split(.FormulaR1C1, Chr(10))  ' arr is the array keep the text
                ub = UBound(arr): st = 1
                ReDim arr1(ub)  ' arr1 is the arry keep the font color
                For j = 0 To ub
                    l = Len(arr(j))
                    arr1(j) = .Characters(Start:=st, Length:=1).Font.Color
                    st = st + l + 1
                Next
                
                .FormulaR1C1 = .FormulaR1C1 & Chr(10) & note
                
                st = 1
                For j = 0 To ub
                    l = Len(arr(j))
                    .Characters(Start:=st, Length:=l).Font.Color = arr1(j)
                    st = st + l + 1
                Next
                
                .Characters(Start:=st, Length:=Len(note)).Font.Color = -16776961    ' red color
                
                .Interior.ThemeColor = xlThemeColorAccent6
                .Interior.TintAndShade = 0.6
            End With
    End Select
End Sub
