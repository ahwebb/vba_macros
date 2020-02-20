Attribute VB_Name = "Occurence_Roll_Up_REFORMAT"

Sub Occurence_Roll_Up_REFORMAT()
Attribute Occurence_Roll_Up_REFORMAT.VB_Description = "Cleans up the Occurence Roll Up to the properly formatted style."
Attribute Occurence_Roll_Up_REFORMAT.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Occurence_Roll_Up_FORMAT Macro
' Cleans up the Occurence Roll Up to the properly formatted style.
'

' Title the Policy Year End column
    Range("B10").Select
    ActiveCell.FormulaR1C1 = "Policy Year End"
    
' Create the Coverage Year column, and paste values to get rid of formula
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "Coverage Year"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(MONTH(DATEVALUE(RC[1]))>=7,CONCAT(YEAR(DATEVALUE(RC[1])),""-"",YEAR(DATEVALUE(RC[1]))+1),CONCAT(YEAR(DATEVALUE(RC[1]))-1,""-"",YEAR(DATEVALUE(RC[1]))))"
    Range("C11").Select
    Selection.AutoFill Destination:=Range("C11:C" & Range("B" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).Select
    Columns("C:C").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
' Title the Member Name and Member Code columns
    Range("H10").Select
    ActiveCell.FormulaR1C1 = "Member Name"
    Range("I10").Select
    ActiveCell.FormulaR1C1 = "Member Code"

' Remove the 0s from the member codes
    Columns("I:I").Select
    Selection.Replace What:="0", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

' Clear column with member names/department codes, populate with only department, name column, copy and paste values to drop formula
    Columns("K:K").Select
    Selection.ClearContents
    Range("K11").Select
    ActiveCell.FormulaR1C1 = _
        "=RIGHT(RIGHT(SUBSTITUTE(RC[-1],"" "",CHAR(9),2),LEN(RC[-1])-FIND(CHAR(9),SUBSTITUTE(RC[-1],"" "",CHAR(9),2),1)+1),LEN(RIGHT(SUBSTITUTE(RC[-1],"" "",CHAR(9),2),LEN(RC[-1])-FIND(CHAR(9),SUBSTITUTE(RC[-1],"" "",CHAR(9),2),1))))"
    Range("K11").Select
    Selection.AutoFill Destination:=Range("K11:K" & Range("B" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).Select
    Columns("K:K").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("K10").Select
    Application.CutCopyMode = False
    Range("K10").Select
    ActiveCell.FormulaR1C1 = "Department"
    
' Delete old Department column redundancy
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft

' Create new column for GG/PO, populate, copy and paste values to drop formula
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    Range("K10").Select
    ActiveCell.FormulaR1C1 = "GG/PO"
    Range("K11").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(ISNUMBER(SEARCH(""LAW ENFORCEMENT"",RC[-1])),COUNTIF('PO_Cities_List.xlsx'!R2C2:R39C2,RC[-2])>0),""PO"",""GG"")"
    Range("K11").Select
    Selection.AutoFill Destination:=Range("K11:K" & Range("B" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).Select
    Columns("K:K").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'   Delete Payment Legal column
    Columns("Y:Y").Select
    Selection.Delete Shift:=xlToLeft
    

'    Columns("AC:AC").Select
'    Selection.Cut

'   Name ending financial columns
    Columns("Y:Y").Select
    Selection.Insert Shift:=xlToRight
    Range("AA10").Select
    ActiveCell.FormulaR1C1 = "Net Paid"
    Range("AB10").Select
    ActiveCell.FormulaR1C1 = "Total Reserves"
    Range("AC10").Select
    ActiveCell.FormulaR1C1 = "Net Incurred"
    Range("AC11").Select

'   Make all font Arial size 11, left and bottom aligned, indent once
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
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.InsertIndent 1
    
'   Format all date columns
    Range("A:B,D:E,O:O").Select
    Range("O1").Activate
    Selection.NumberFormat = "mm/dd/yy;@"

'   Set financial columns to accounting format
    Columns("Q:AD").Select
    Selection.Style = "Currency"
    Selection.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"

'   Nudge all the text flagging
    Cells.Select
    Selection.Replace What:="=T(""", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:=""")", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
'   Create Cov Yr/Mbr column, copy and paste to remove formulas
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
    Range("J10").Select
    ActiveCell.FormulaR1C1 = "Cov Yr/Mbr"
    Range("J11").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-7]&RC[-1]"
    Range("J11").Select
    Selection.AutoFill Destination:=Range("J11:J" & Range("B" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).Select
    Columns("J:J").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
'   Format column widths
    Columns("A:A").ColumnWidth = 11
    Columns("B:B").ColumnWidth = 11
    Columns("C:C").ColumnWidth = 12
    Columns("D:D").ColumnWidth = 10
    Columns("E:E").ColumnWidth = 10
    Columns("F:F").ColumnWidth = 10.5
    Columns("G:G").ColumnWidth = 10
    Columns("H:H").ColumnWidth = 10
    Columns("I:I").ColumnWidth = 10
    Columns("J:J").EntireColumn.AutoFit
    Columns("K:K").ColumnWidth = 14
    Columns("L:L").ColumnWidth = 9
    Columns("M:M").ColumnWidth = 64.5
    Columns("N:N").ColumnWidth = 10
    Columns("O:O").ColumnWidth = 12
    Columns("P:P").EntireColumn.AutoFit
    Columns("Q:Q").EntireColumn.AutoFit
    Columns("R:R").EntireColumn.AutoFit
    Columns("T:T").EntireColumn.AutoFit
    Columns("U:U").EntireColumn.AutoFit
    Columns("V:V").EntireColumn.AutoFit
    Columns("W:W").EntireColumn.AutoFit
    Columns("Y:Y").EntireColumn.AutoFit
    Columns("AA:AA").EntireColumn.AutoFit
    Columns("AB:AB").EntireColumn.AutoFit
    Columns("AC:AC").EntireColumn.AutoFit
    Columns("AD:AD").ColumnWidth = 16
    Columns("AE:AE").ColumnWidth = 16
    
'   Remove empty column
    Columns("Z:Z").Select
    Selection.Delete Shift:=xlToLeft
    
'   Remove excess rows at the beginning of the spreadsheet
    Range("A10").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Cut
    Range("A1").Select
    ActiveSheet.Paste
    
' Fill the header row, align, and bold text
    Rows("1:1").RowHeight = 45
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    
'   Freeze active window
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    
'   Display message box to confirm macro is complete
    MsgBox "Reformatting complete!"
End Sub
