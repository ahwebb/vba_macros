Attribute VB_Name = "Injury_Year_Detail_REFORMAT"

Sub Injury_Year_Detail_REFORMAT()
Attribute Injury_Year_Detail_REFORMAT.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Injury_Year_Detail_REFORMAT Macro
' Cleans up the Injury Year Detail report to the properly formatted style.

Dim StartTime As Double
Dim MinutesElapsed As String

    StartTime = Timer

' Display gridlines
    ActiveWindow.DisplayGridlines = True
    
' Create column to store Date of Loss as text
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    Range("H7").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""mm/dd/yy"")"
    Selection.AutoFill Destination:=Range("H7:H" & Range("B" & Rows.Count).End(xlUp).Row)
    
' Create Coverage Year column, and paste to remove formula
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    Range("I6").Select
    ActiveCell.FormulaR1C1 = "Coverage Year"
    Range("I7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(MONTH(DATEVALUE(RC[-1]))>=7,CONCAT(YEAR(DATEVALUE(RC[-1])),""-"",YEAR(DATEVALUE(RC[-1]))+1),CONCAT(YEAR(DATEVALUE(RC[-1]))-1,""-"",YEAR(DATEVALUE(RC[-1]))))"
    Selection.AutoFill Destination:=Range("I7:I" & Range("B" & Rows.Count).End(xlUp).Row)
    Columns("I:I").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
' Create Limit column with accurate values, paste to remove formula
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    Range("L6").Select
    ActiveCell.FormulaR1C1 = "Limit"
    Range("L7").Select
    ActiveCell.FormulaR1C1 = _
        "=IFS(YEAR(DATEVALUE(RC[-4]))<1983,250000,AND(YEAR(DATEVALUE(RC[-4]))=1983,MONTH(DATEVALUE(RC[-4]))<=3),250000,AND(YEAR(DATEVALUE(RC[-4]))=1983,MONTH(DATEVALUE(RC[-4]))>3),100000,YEAR(DATEVALUE(RC[-4]))=1984,100000,AND(YEAR(DATEVALUE(RC[-4]))=1985,MONTH(DATEVALUE(RC[-4]))<=6),100000,AND(YEAR(DATEVALUE(RC[-4]))=1985,MONTH(DATEVALUE(RC[-4]))>6),400000,AND(YEAR(DATEVALUE(RC[-4]))=1986,MONTH(DATEVALUE(RC[-4]))<=6),400000,AND(YEAR(DATEVALUE(RC[-4]))=1986,MONTH(DATEVALUE(RC[-4]))>6),500000,AND(YEAR(DATEVALUE(RC[-4]))>=1987,YEAR(DATEVALUE(RC[-4]))<=2001),500000,AND(YEAR(DATEVALUE(RC[-4]))=2002,MONTH(DATEVALUE(RC[-4]))<=6),500000,AND(YEAR(DATEVALUE(RC[-4]))=2002,MONTH(DATEVALUE(RC[-4]))>6),2000000,YEAR(DATEVALUE(RC[-4]))>=2003,2000000)"
    Selection.AutoFill Destination:=Range("L7:L" & Range("B" & Rows.Count).End(xlUp).Row)
    Columns("L:L").Select
    Selection.Style = "Currency"
    Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
' Delete old Limit and Coverage Year columns
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    
' Convert dates to mm/dd/yy format
    Columns("K:Q").Select
    Selection.NumberFormat = "mm/dd/yy;@"
    Columns("G:G").Select
    Selection.NumberFormat = "mm/dd/yy;@"
    
' Create GG/PS column
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    Range("AB6").Select
    ActiveCell.FormulaR1C1 = "GG/PS"
    Range("AB7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(OR(ISNUMBER(SEARCH(""7720"",RC[-1])),ISNUMBER(SEARCH(""7721"",RC[-1])),ISNUMBER(SEARCH(""7706"",RC[-1])),ISNUMBER(SEARCH(""7707"",RC[-1]))),COUNTIF('PS_Cities_List.xlsx'!R2C2:R50C2,RC[-26])>0),""PS"",""GG"")"
    Selection.AutoFill Destination:=Range("AB7:AB" & Range("B" & Rows.Count).End(xlUp).Row)
    
' Copy and paste GG/PS values to get rid of formula
    Columns("AB:AB").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
' Move Record Only Transactions and 4850 Diff (Voucher) to the end of the sheet
    Columns("AN:AO").Select
    Selection.Cut
    Columns("AZ:BA").Select
    Selection.Insert Shift:=xlToRight
    
' Unmerge cells and clear formatting
    Range("AU1:AW3").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    Selection.UnMerge
    Range("AU1:AU3").Select
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    
' Place a borders to segment Total Reserves, Total Paid, and Total Incurred
    Range("AM:AM, AS:AS, AW:AW").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
' Rename columns
    Range("AS6").Select
    ActiveCell.FormulaR1C1 = "Total Paid"
    Range("AV6").Select
    ActiveCell.FormulaR1C1 = "Net Paid"
    Range("AZ6").Select
    ActiveCell.FormulaR1C1 = "4850 Diff Reserves 7/1/09 & AFTER"
    Range("BA6").Select
    ActiveCell.FormulaR1C1 = "Gross Paid"
    Range("BB6").Select
    ActiveCell.FormulaR1C1 = "Gross Reserved"
    Range("BC6").Select
    ActiveCell.FormulaR1C1 = "Gross Incurred"
    
' Change formatting of created columns
    Columns("AZ:BC").Select
    With Selection.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("AZ6:BC6").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .MergeCells = False
    End With
    
' Populate 4850 Diff Reserves 6/30/09 & PRIOR column, and delete old 4850 Diff Reserves column
    Columns("AJ:AJ").Select
    Selection.Insert Shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    Range("AJ7").Select
    ActiveCell.FormulaR1C1 = "=IF(DATEVALUE(RC[-28])<39995,RC[-1],0)"
    Selection.AutoFill Destination:=Range("AJ7:AJ" & Range("B" & Rows.Count).End(xlUp).Row)
    Range("AJ6").Select
    ActiveCell.FormulaR1C1 = "4850 Diff Reserves 6/30/09 & PRIOR"
    Columns("AJ:AJ").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Columns("AI:AI").Select
    Selection.Delete Shift:=xlToLeft
    
' Populate 4850 Diff Reserves 7/1/09 & AFTER column
    Columns("AZ:AZ").EntireColumn.AutoFit
    Range("AZ7").Select
    ActiveCell.FormulaR1C1 = "=IF(DATEVALUE(RC[-44])<39995,0,RC[-16])"
    Selection.AutoFill Destination:=Range("AZ7:AZ" & Range("B" & Rows.Count).End(xlUp).Row)
    Columns("AZ:AZ").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
' Populate Gross Paid column
    Range("BA7").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2],RC[-8])"
    Selection.AutoFill Destination:=Range("BA7:BA" & Range("B" & Rows.Count).End(xlUp).Row)
    Columns("BA:BA").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
' Populate Gross Reserved column
    Range("BB7").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2],RC[-15])"
    Selection.AutoFill Destination:=Range("BB7:BB" & Range("B" & Rows.Count).End(xlUp).Row)
    Columns("BB:BB").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
' Populate Gross Incurred column
    Range("BC7").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-1],RC[-2])"
    Selection.AutoFill Destination:=Range("BC7:BC" & Range("B" & Rows.Count).End(xlUp).Row)
    Columns("BC:BC").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
' Clean up data storage column
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    
' Format entire sheet: left align, bottom align, indent once
    Cells.Select
    With Selection
        .HorizontalAlignment = xlLeft
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
    End With
    Selection.InsertIndent 1
    
' Set financial numbers to accounting format
    Columns("AG:BB").Select
    Selection.Style = "currency"
    Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    
' Increase font size and resize columns
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 11
    End With
    Selection.ColumnWidth = 60
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    Columns("AG:BB").Select
    Selection.ColumnWidth = 21
    Columns("A:A").ColumnWidth = 17
    Columns("C:C").ColumnWidth = 21
    Columns("AI:AI").EntireColumn.AutoFit
    Columns("AJ:AJ").ColumnWidth = 18.57
    Columns("AJ:AJ").EntireColumn.AutoFit
    Columns("AK:AK").EntireColumn.AutoFit
    Columns("AL:AL").ColumnWidth = 16.43
    Columns("AL:AL").EntireColumn.AutoFit
    Columns("AM:AM").ColumnWidth = 16.86
    Columns("AM:AM").EntireColumn.AutoFit
    Columns("AN:AN").ColumnWidth = 13.71
    Columns("AN:AN").EntireColumn.AutoFit
    Columns("AO:AO").EntireColumn.AutoFit
    Columns("AP:AP").EntireColumn.AutoFit
    Columns("AQ:AQ").ColumnWidth = 21.14
    Columns("AQ:AQ").EntireColumn.AutoFit
    Columns("AR:AR").EntireColumn.AutoFit
    Columns("AS:AS").EntireColumn.AutoFit
    Columns("AT:AT").ColumnWidth = 18.57
    Columns("AT:AT").EntireColumn.AutoFit
    Columns("AU:AU").EntireColumn.AutoFit
    Columns("AV:AV").EntireColumn.AutoFit
    Columns("AW:AW").EntireColumn.AutoFit
    Columns("AX:AX").ColumnWidth = 18.29
    Columns("AX:AX").EntireColumn.AutoFit
    Columns("AZ:AZ").EntireColumn.AutoFit
    Columns("BA:BA").ColumnWidth = 17.71
    Columns("BA:BA").EntireColumn.AutoFit
    Columns("BB:BB").EntireColumn.AutoFit
    Columns("F:F").ColumnWidth = 6.14
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").ColumnWidth = 10.43
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").ColumnWidth = 15.43
    Columns("H:H").EntireColumn.AutoFit
    Columns("J:J").ColumnWidth = 12.57
    Columns("J:J").EntireColumn.AutoFit
    Columns("K:K").ColumnWidth = 15.86
    Columns("K:K").EntireColumn.AutoFit
    Columns("L:L").ColumnWidth = 16.14
    Columns("L:L").EntireColumn.AutoFit
    Columns("M:M").ColumnWidth = 11.57
    Columns("M:M").EntireColumn.AutoFit
    Columns("N:N").ColumnWidth = 9.86
    Columns("N:N").EntireColumn.AutoFit
    Columns("O:O").ColumnWidth = 11
    Columns("O:O").EntireColumn.AutoFit
    Columns("P:P").ColumnWidth = 12
    Columns("P:P").EntireColumn.AutoFit
    Columns("Q:Q").ColumnWidth = 13
    Columns("Q:Q").EntireColumn.AutoFit
    Columns("R:R").ColumnWidth = 11.29
    Columns("R:R").EntireColumn.AutoFit
    
' Remove excess rows at the beginning of the spreadsheet
    Range("A6").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Cut
    Range("A1").Select
    ActiveSheet.Paste
    
' Fill header row
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    Range("A1").Select

' Freeze active window
    Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    
' Disable word wrap except in header row
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
    End With
        
MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    
' Display message box on completion
    MsgBox "Reformatting was completed successfully in  " & MinutesElapsed & "!", vbInformation

End Sub
