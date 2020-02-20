Attribute VB_Name = "Occurence_Roll_Up_REFORMAT"
Sub Occurence_Roll_Up_REFORMAT()
Attribute Occurence_Roll_Up_REFORMAT.VB_Description = "Cleans up the Occurence Roll Up to the properly formatted style."
Attribute Occurence_Roll_Up_REFORMAT.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Occurence_Roll_Up_FORMAT Macro
' Cleans up the Occurence Roll Up to the properly formatted style.
'

'
    Rows("10:10").RowHeight = 45
    Rows("10:10").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B10").Select
    ActiveCell.FormulaR1C1 = "Policy Year End"
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "Coverage Year"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(MONTH(DATEVALUE(RC[1]))>=7,CONCAT(YEAR(DATEVALUE(RC[1])),""-"",YEAR(DATEVALUE(RC[1]))+1),CONCAT(YEAR(DATEVALUE(RC[1]))-1,""-"",YEAR(DATEVALUE(RC[1]))))"
    Range("C11").Select
    Selection.AutoFill Destination:=Range("C11:C50000")
    Range("C11:C50000").Select
    ActiveWindow.SmallScroll Down:=135
    ActiveWindow.ScrollRow = 137
    ActiveWindow.ScrollRow = 140
    ActiveWindow.ScrollRow = 146
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 158
    ActiveWindow.ScrollRow = 168
    ActiveWindow.ScrollRow = 171
    ActiveWindow.ScrollRow = 180
    ActiveWindow.ScrollRow = 183
    ActiveWindow.ScrollRow = 192
    ActiveWindow.ScrollRow = 199
    ActiveWindow.ScrollRow = 202
    ActiveWindow.ScrollRow = 208
    ActiveWindow.ScrollRow = 214
    ActiveWindow.ScrollRow = 223
    ActiveWindow.ScrollRow = 226
    ActiveWindow.ScrollRow = 233
    ActiveWindow.ScrollRow = 236
    ActiveWindow.ScrollRow = 257
    ActiveWindow.ScrollRow = 267
    ActiveWindow.ScrollRow = 295
    ActiveWindow.ScrollRow = 313
    ActiveWindow.ScrollRow = 325
    ActiveWindow.ScrollRow = 335
    ActiveWindow.ScrollRow = 338
    ActiveWindow.ScrollRow = 350
    ActiveWindow.ScrollRow = 356
    ActiveWindow.ScrollRow = 390
    ActiveWindow.ScrollRow = 412
    ActiveWindow.ScrollRow = 428
    ActiveWindow.ScrollRow = 434
    ActiveWindow.ScrollRow = 449
    ActiveWindow.ScrollRow = 468
    ActiveWindow.ScrollRow = 477
    ActiveWindow.ScrollRow = 499
    ActiveWindow.ScrollRow = 505
    ActiveWindow.ScrollRow = 517
    ActiveWindow.ScrollRow = 520
    ActiveWindow.ScrollRow = 533
    ActiveWindow.ScrollRow = 536
    ActiveWindow.ScrollRow = 545
    ActiveWindow.ScrollRow = 551
    ActiveWindow.ScrollRow = 555
    ActiveWindow.ScrollRow = 561
    ActiveWindow.ScrollRow = 570
    ActiveWindow.ScrollRow = 582
    ActiveWindow.ScrollRow = 589
    ActiveWindow.ScrollRow = 607
    ActiveWindow.ScrollRow = 626
    ActiveWindow.ScrollRow = 629
    ActiveWindow.ScrollRow = 638
    ActiveWindow.ScrollRow = 641
    ActiveWindow.ScrollRow = 650
    ActiveWindow.ScrollRow = 660
    ActiveWindow.ScrollRow = 672
    ActiveWindow.ScrollRow = 675
    ActiveWindow.ScrollRow = 678
    ActiveWindow.ScrollRow = 681
    ActiveWindow.ScrollRow = 688
    ActiveWindow.ScrollRow = 694
    ActiveWindow.ScrollRow = 700
    ActiveWindow.ScrollRow = 703
    ActiveWindow.ScrollRow = 709
    ActiveWindow.ScrollRow = 715
    ActiveWindow.ScrollRow = 722
    ActiveWindow.ScrollRow = 728
    ActiveWindow.ScrollRow = 731
    ActiveWindow.ScrollRow = 746
    ActiveWindow.ScrollRow = 753
    ActiveWindow.ScrollRow = 771
    ActiveWindow.ScrollRow = 780
    ActiveWindow.ScrollRow = 811
    ActiveWindow.ScrollRow = 824
    ActiveWindow.ScrollRow = 845
    ActiveWindow.ScrollRow = 852
    ActiveWindow.ScrollRow = 867
    ActiveWindow.ScrollRow = 870
    ActiveWindow.ScrollRow = 879
    ActiveWindow.ScrollRow = 889
    ActiveWindow.ScrollRow = 898
    ActiveWindow.ScrollRow = 907
    ActiveWindow.ScrollRow = 920
    ActiveWindow.ScrollRow = 926
    ActiveWindow.ScrollRow = 929
    ActiveWindow.ScrollRow = 932
    ActiveWindow.ScrollRow = 938
    ActiveWindow.ScrollRow = 951
    ActiveWindow.ScrollRow = 954
    ActiveWindow.ScrollRow = 963
    ActiveWindow.ScrollRow = 969
    ActiveWindow.ScrollRow = 979
    ActiveWindow.ScrollRow = 982
    ActiveWindow.ScrollRow = 991
    ActiveWindow.ScrollRow = 997
    ActiveWindow.ScrollRow = 1009
    ActiveWindow.ScrollRow = 1013
    ActiveWindow.ScrollRow = 1040
    ActiveWindow.ScrollRow = 1053
    ActiveWindow.ScrollRow = 1078
    ActiveWindow.ScrollRow = 1090
    ActiveWindow.ScrollRow = 1109
    ActiveWindow.ScrollRow = 1112
    ActiveWindow.ScrollRow = 1115
    ActiveWindow.ScrollRow = 1118
    ActiveWindow.ScrollRow = 1121
    ActiveWindow.ScrollRow = 1124
    ActiveWindow.ScrollRow = 1127
    ActiveWindow.ScrollRow = 1133
    ActiveWindow.ScrollRow = 1143
    ActiveWindow.ScrollRow = 1152
    ActiveWindow.ScrollRow = 1161
    ActiveWindow.ScrollRow = 1167
    ActiveWindow.ScrollRow = 1177
    ActiveWindow.ScrollRow = 1183
    ActiveWindow.ScrollRow = 1192
    ActiveWindow.ScrollRow = 1195
    ActiveWindow.ScrollRow = 1204
    ActiveWindow.ScrollRow = 1211
    ActiveWindow.ScrollRow = 1223
    ActiveWindow.ScrollRow = 1229
    ActiveWindow.ScrollRow = 1245
    ActiveWindow.ScrollRow = 1260
    ActiveWindow.ScrollRow = 1276
    ActiveWindow.ScrollRow = 1291
    ActiveWindow.ScrollRow = 1322
    ActiveWindow.ScrollRow = 1350
    ActiveWindow.ScrollRow = 1396
    ActiveWindow.ScrollRow = 1415
    ActiveWindow.ScrollRow = 1430
    ActiveWindow.ScrollRow = 1492
    ActiveWindow.ScrollRow = 1508
    ActiveWindow.ScrollRow = 1539
    ActiveWindow.ScrollRow = 1548
    ActiveWindow.ScrollRow = 1563
    ActiveWindow.ScrollRow = 1567
    ActiveWindow.ScrollRow = 1573
    ActiveWindow.ScrollRow = 1579
    ActiveWindow.ScrollRow = 1610
    ActiveWindow.ScrollRow = 1641
    ActiveWindow.ScrollRow = 1681
    ActiveWindow.ScrollRow = 1712
    ActiveWindow.ScrollRow = 1727
    ActiveWindow.ScrollRow = 1755
    ActiveWindow.ScrollRow = 1771
    ActiveWindow.ScrollRow = 1796
    ActiveWindow.ScrollRow = 1817
    ActiveWindow.ScrollRow = 1839
    ActiveWindow.ScrollRow = 1842
    ActiveWindow.ScrollRow = 1845
    ActiveWindow.ScrollRow = 1851
    ActiveWindow.ScrollRow = 1854
    ActiveWindow.ScrollRow = 1861
    ActiveWindow.ScrollRow = 1864
    ActiveWindow.ScrollRow = 1867
    ActiveWindow.ScrollRow = 1870
    ActiveWindow.ScrollRow = 1873
    ActiveWindow.ScrollRow = 1876
    ActiveWindow.ScrollRow = 1879
    ActiveWindow.ScrollRow = 1882
    ActiveWindow.ScrollRow = 1885
    ActiveWindow.ScrollRow = 1888
    ActiveWindow.ScrollRow = 1895
    ActiveWindow.ScrollRow = 1898
    ActiveWindow.ScrollRow = 1901
    ActiveWindow.ScrollRow = 1904
    ActiveWindow.ScrollRow = 1907
    ActiveWindow.ScrollRow = 1910
    ActiveWindow.ScrollRow = 1913
    ActiveWindow.ScrollRow = 1916
    ActiveWindow.ScrollRow = 1926
    ActiveWindow.ScrollRow = 1932
    ActiveWindow.ScrollRow = 1935
    ActiveWindow.ScrollRow = 1938
    ActiveWindow.ScrollRow = 1941
    ActiveWindow.ScrollRow = 1944
    ActiveWindow.ScrollRow = 1947
    ActiveWindow.ScrollRow = 1950
    ActiveWindow.ScrollRow = 1957
    ActiveWindow.ScrollRow = 1960
    ActiveWindow.ScrollRow = 1966
    ActiveWindow.ScrollRow = 1969
    ActiveWindow.ScrollRow = 1975
    ActiveWindow.ScrollRow = 1978
    ActiveWindow.ScrollRow = 1984
    ActiveWindow.ScrollRow = 1991
    ActiveWindow.ScrollRow = 1994
    ActiveWindow.ScrollRow = 2000
    ActiveWindow.ScrollRow = 2003
    ActiveWindow.ScrollRow = 2012
    ActiveWindow.ScrollRow = 2015
    ActiveWindow.ScrollRow = 2028
    ActiveWindow.ScrollRow = 2031
    ActiveWindow.ScrollRow = 2037
    ActiveWindow.ScrollRow = 2043
    ActiveWindow.ScrollRow = 2046
    ActiveWindow.ScrollRow = 2049
    ActiveWindow.ScrollRow = 2056
    ActiveWindow.ScrollRow = 2059
    ActiveWindow.ScrollRow = 2065
    ActiveWindow.ScrollRow = 2068
    ActiveWindow.ScrollRow = 2071
    ActiveWindow.ScrollRow = 2077
    ActiveWindow.ScrollRow = 2080
    ActiveWindow.ScrollRow = 2083
    ActiveWindow.ScrollRow = 2093
    ActiveWindow.ScrollRow = 2099
    ActiveWindow.ScrollRow = 2102
    ActiveWindow.ScrollRow = 2105
    ActiveWindow.ScrollRow = 2108
    ActiveWindow.ScrollRow = 2111
    ActiveWindow.ScrollRow = 2114
    ActiveWindow.ScrollRow = 2121
    ActiveWindow.ScrollRow = 2124
    ActiveWindow.ScrollRow = 2130
    ActiveWindow.ScrollRow = 2136
    ActiveWindow.ScrollRow = 2139
    ActiveWindow.ScrollRow = 2148
    ActiveWindow.ScrollRow = 2152
    ActiveWindow.ScrollRow = 2148
    ActiveWindow.ScrollRow = 2096
    ActiveWindow.ScrollRow = 1731
    ActiveWindow.ScrollRow = 1539
    ActiveWindow.ScrollRow = 1090
    ActiveWindow.ScrollRow = 938
    ActiveWindow.ScrollRow = 635
    ActiveWindow.ScrollRow = 561
    ActiveWindow.ScrollRow = 335
    ActiveWindow.ScrollRow = 273
    ActiveWindow.ScrollRow = 165
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 1
    Range("H10").Select
    ActiveCell.FormulaR1C1 = "=T(""Member Name"")"
    Range("I10").Select
    ActiveCell.FormulaR1C1 = "Member Code"
    Columns("I:I").Select
    Selection.Replace What:="0", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Columns("K:K").Select
    Selection.ClearContents
    Range("K11").Select
    ActiveCell.FormulaR1C1 = _
        "=RIGHT(RIGHT(SUBSTITUTE(RC[-1],"" "",CHAR(9),2),LEN(RC[-1])-FIND(CHAR(9),SUBSTITUTE(RC[-1],"" "",CHAR(9),2),1)+1),LEN(RIGHT(SUBSTITUTE(RC[-1],"" "",CHAR(9),2),LEN(RC[-1])-FIND(CHAR(9),SUBSTITUTE(RC[-1],"" "",CHAR(9),2),1))))"
    Range("K11").Select
    Selection.AutoFill Destination:=Range("K11:K50000")
    Range("K11:K50000").Select
    Columns("K:K").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K10").Select
    Application.CutCopyMode = False
    Range("K10").Select
    ActiveCell.FormulaR1C1 = "Department"
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("K10").Select
    ActiveCell.FormulaR1C1 = "GG/PO"
    Range("K11").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(ISNUMBER(SEARCH(""LAW ENFORCEMENT"",RC[-1])),COUNTIF('PO Cities List.xlsx'!R2C2:R39C2,RC[-2])>0),""PO"",""GG"")"
    Range("K11").Select
    Selection.AutoFill Destination:=Range("K11:K50000")
    Range("K11:K50000").Select
    Columns("K:K").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWindow.SmallScroll ToRight:=6
    Columns("Y:Y").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AC:AC").Select
    Selection.Cut
    Columns("Y:Y").Select
    Selection.Insert Shift:=xlToRight
    Range("AA10").Select
    ActiveCell.FormulaR1C1 = "=T(""Net Paid"")"
    Range("AB10").Select
    ActiveCell.FormulaR1C1 = "=T(""Total Reserves"")"
    Range("AC10").Select
    ActiveCell.FormulaR1C1 = "=T(""Net Incurred"")"
    Range("AC11").Select
    ActiveWindow.SmallScroll ToRight:=-15
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("A:B,D:E,O:O").Select
    Range("O1").Activate
    Selection.NumberFormat = "m/d/yy;@"
    Selection.NumberFormat = "mm/dd/yy;@"
    ActiveWindow.SmallScroll ToRight:=10
    Columns("Q:AC").Select
    Selection.Style = "Currency"
    Selection.NumberFormat = "_($* #,##0.0_);_($* (#,##0.0);_($* ""-""??_);_(@_)"
    Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Cells.Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.InsertIndent 1
    Cells.Select
    Selection.Replace What:="=T(""", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:=""")", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.ColumnWidth = 11.57
    ActiveWindow.SmallScroll ToRight:=12
    ActiveWindow.SmallScroll Down:=12
    ActiveWindow.SmallScroll ToRight:=-15
    ActiveWindow.SmallScroll Down:=-12
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("J10").Select
    ActiveCell.FormulaR1C1 = "Cov Yr/Mbr"
    Range("J11").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-7]&RC[-1]"
    Range("J11").Select
    Selection.AutoFill Destination:=Range("J11:J50000")
    Range("J11:J50000").Select
    Columns("J:J").ColumnWidth = 15.43
    Columns("J:J").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("I10").Select
    ActiveWindow.SmallScroll ToRight:=12
    ActiveWindow.SmallScroll Down:=3
    Columns("R:AD").Select
    Range("R4").Activate
    Selection.ColumnWidth = 15
    ActiveWindow.SmallScroll ToRight:=3
    Range("AD30").Select
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=-12
    Range("A11").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.SmallScroll Down:=-9
    Range("K:K,M:M").Select
    Selection.ColumnWidth = 15
    ActiveSheet.Cells(1, 1).Select
    MsgBox "Reformatting complete!"
End Sub
