Option Explicit

Dim DataWorkbook As Workbook
Dim DataSheet As Worksheet
Dim RotCoordLookupTable As Workbook
Dim EPALookupTable As Workbook
Dim ExtractTable As ListObject

Sub Generate_Rot_Coord_Block_Report()
    Application.ScreenUpdating = False
    Call Format_Block_Report
    Application.ScreenUpdating = True
    Range("A1").Select
    Application.CutCopyMode = False

End Sub

Sub Format_Block_Report()

    Dim fso As Scripting.FileSystemObject
    Dim fd As FileDialog
    Dim fileWasChosen As Boolean
    Dim BlockReportCopy As String
    Dim firstYear As String
    Dim secondYear As String
    Dim newCol As Range
    Dim newHeader As Range
    Dim formulaString As String
    
    Set fso = New Scripting.FileSystemObject
    
    MsgBox "Choose a Block Report to generate from..."
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    fd.Filters.Clear
    fd.Filters.Add "CSV Files", "*.csv, *.xl*"
    fd.FilterIndex = 1
    
    fd.AllowMultiSelect = False
    fd.InitialFileName = "N:\DOM1\Post Graduate Program"
    fd.Title = "Choose an Extract Data File"
    
    fileWasChosen = fd.Show
    
    If Not fileWasChosen Then
        MsgBox "You didn't choose a Block Report File. Report generation was terminated."
        End
    End If
    
    ' Creates a copy of the Block Report data to preserve the original. Stored in a variable for reference when
    ' opening the new workbook
    BlockReportCopy = Left(fd.SelectedItems(1), InStrRev(fd.SelectedItems(1), ".") - 1) & " copy." & _
        Right(fd.SelectedItems(1), Len(fd.SelectedItems(1)) - InStrRev(fd.SelectedItems(1), "."))
    
    fso.CopyFile fd.SelectedItems(1), BlockReportCopy
    
    Set DataWorkbook = Workbooks.Open(Filename:=BlockReportCopy)
    
    ' For naming the file.
    ' If current month is before July, then the date format is
    ' PreviousYear - CurrentYear
    ' Else, it is CurrentYear - NextYear
    If Month(Date) < 7 Then
        firstYear = Str(Year(Date) - 1)
        secondYear = Right(Str(Year(Date)), 2)
    Else
        firstYear = Str(Year(Date))
        secondYear = Right(Str(Year(Date) + 1), 2)
    End If
    
    ' Rename the workbook copy with the appropriate division name and date range
    DataWorkbook.SaveAs DataWorkbook.Path & "\" & "Block " & DataWorkbook.ActiveSheet.Range("A3").Value & _
        " Rotation Coordinator " & firstYear & "-" & secondYear & ".xlsx", 51
    Set DataSheet = DataWorkbook.Worksheets(1)
    DataSheet.Name = "OriginalSheet"
    
    ' Create a table from the data for ease of manipulation - allows use of column
    ' header names for flexibility
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes).Name = _
        "ExtractTable"
    Set ExtractTable = ActiveSheet.ListObjects("ExtractTable")

    ' Insert Column for Concatenation of first 3 columns
    ExtractTable.ListColumns("PGY1s").Range.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "PeriodRotationHospital"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=[@Period]&[@Rotation]&[@Hospital]"
    ExtractTable.ListColumns("PeriodRotationHospital").Range.EntireColumn.AutoFit
    
    ExtractTable.Range.Cells.Select
    ' Replace all NULLS
    Selection.Replace What:="NULL", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False
    ' Break names onto new line by space-comma
    Selection.Replace What:=" ,", Replacement:="" & Chr(10) & "", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False
    ' Break names onto new line by comma
    Selection.Replace What:=",", Replacement:="" & Chr(10) & "", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False
        
    ' Insert VLOOKUP Columns
    ' Obtain the Rotation Coordinator Contact info lookup table
     
    MsgBox "Choose a Rotation Coordinator Contact Info VLOOKUP Table for the report."
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xl*"
    fd.FilterIndex = 1
    
    fd.AllowMultiSelect = False
    fd.InitialFileName = "N:\DOM1\Post Graduate Program\IM\Competence by Design - IM\Communications\Rotation Coordinator"
    fd.Title = "Choose a Rotation Coordinator Contact Info Lookup Table"
    
    fileWasChosen = fd.Show
    
    If Not fileWasChosen Then
        MsgBox "You didn't choose a Rotation Coordinator Contact Info Lookup Table. Report generation was terminated."
        End
    End If
    
    Set RotCoordLookupTable = Workbooks.Open(Filename:=fd.SelectedItems(1))
    DataWorkbook.Activate
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "Rotation Coordinator"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@Rotation]&" & Chr(34) & " - " & Chr(34) & "&[@Hospital],'[" & RotCoordLookupTable.Name & "]RC'!$C:$H,2,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    ExtractTable.ListColumns("Rotation Coordinator").Range.EntireColumn.AutoFit
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "RC First Name"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@Rotation]&" & Chr(34) & " - " & Chr(34) & "&[@Hospital],'[" & RotCoordLookupTable.Name & "]RC'!$C:$H,3,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    ExtractTable.ListColumns("RC First Name").Range.EntireColumn.AutoFit
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "RC Email"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@Rotation]&" & Chr(34) & " - " & Chr(34) & "&[@Hospital],'[" & RotCoordLookupTable.Name & "]RC'!$C:$H,4,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    ExtractTable.ListColumns("RC Email").Range.EntireColumn.AutoFit
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "Assistant"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@Rotation]&" & Chr(34) & " - " & Chr(34) & "&[@Hospital],'[" & RotCoordLookupTable.Name & "]RC'!$C:$H,5,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    ExtractTable.ListColumns("Assistant").Range.EntireColumn.AutoFit
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "Assistant Email"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@Rotation]&" & Chr(34) & " - " & Chr(34) & "&[@Hospital],'[" & RotCoordLookupTable.Name & "]RC'!$C:$H,6,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    ExtractTable.ListColumns("Assistant Email").Range.EntireColumn.AutoFit
    
    ' Obtain the EPA and Rotation Card Lookup table
     
    MsgBox "Choose an EPA and Rotation Card VLOOKUP Table for the report."
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xl*"
    fd.FilterIndex = 1
    
    fd.AllowMultiSelect = False
    fd.InitialFileName = "N:\DOM1\Post Graduate Program\IM\Competence by Design - IM\Communications\Rotation Coordinator"
    fd.Title = "Choose an EPA and Rotation Card Lookup Table"
    
    fileWasChosen = fd.Show
    
    If Not fileWasChosen Then
        MsgBox "You didn't choose a Rotation Coordinator Contact Info Lookup Table. Report generation was terminated."
        End
    End If
    
    Set EPALookupTable = Workbooks.Open(Filename:=fd.SelectedItems(1))
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY1 Priority (Highest*) when you can Optional"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@PeriodRotationHospital],'[" & EPALookupTable.Name & "]Sheet1'!$D:$N,4,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY1 Always do when you can"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@PeriodRotationHospital],'[" & EPALookupTable.Name & "]Sheet1'!$D:$N,5,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY2 Priority (Highest*) when you can Optional"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@PeriodRotationHospital],'[" & EPALookupTable.Name & "]Sheet1'!$D:$N,7,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY2 Always do when you can"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@PeriodRotationHospital],'[" & EPALookupTable.Name & "]Sheet1'!$D:$N,8,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY1 Rotation Cards "
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@PeriodRotationHospital],'[" & EPALookupTable.Name & "]Sheet1'!$D:$N,10,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY2 Rotation Cards "
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@PeriodRotationHospital],'[" & EPALookupTable.Name & "]Sheet1'!$D:$N,11,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    
    '****************************
    '   Copy PGY1s to New Sheet
    '****************************
    Dim PGY1ColNum As Integer
    Dim PGY2ColNum As Integer
    
    PGY1ColNum = ExtractTable.ListColumns("PGY1s").Range.Column
    PGY2ColNum = ExtractTable.ListColumns("PGY2s").Range.Column
    
    DataWorkbook.Worksheets.Add After:=DataWorkbook.Worksheets("OriginalSheet")
    DataWorkbook.Worksheets("Sheet1").Name = "PGY1"
    
    ' Filter PGY1s to not show blanks
    ExtractTable.Range.AutoFilter Field:=PGY1ColNum, Criteria1 _
        :="<>"
    ' Filter PGY2s to only show blanks
    ExtractTable.Range.AutoFilter Field:=PGY2ColNum, Criteria1 _
        :="="
        
    ExtractTable.Range.SpecialCells(xlCellTypeVisible).Copy
    
    DataWorkbook.Worksheets("PGY1").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    
    '****************************
    '   Copy PGY2s to New Sheet
    '****************************
    
    DataWorkbook.Worksheets.Add After:=DataWorkbook.Worksheets("PGY1")
    DataWorkbook.Worksheets("Sheet2").Name = "PGY2"
    
    'Filter PGY2s to not show blanks
    ExtractTable.Range.AutoFilter Field:=PGY2ColNum, Criteria1 _
        :="<>"
    ' Filter PGY1s to only show blanks
    ExtractTable.Range.AutoFilter Field:=PGY1ColNum, Criteria1 _
        :="="
        
    ExtractTable.Range.SpecialCells(xlCellTypeVisible).Copy
    
    DataWorkbook.Worksheets("PGY2").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    '****************************
    '   Copy PGY1&2s to New Sheet
    '****************************
    
    DataWorkbook.Worksheets.Add After:=DataWorkbook.Worksheets("PGY2")
    DataWorkbook.Worksheets("Sheet3").Name = "PGY1&2"
    
    ' Filter PGY1s to not show blanks
    ExtractTable.Range.AutoFilter Field:=PGY1ColNum, Criteria1 _
        :="<>"
    ' Filter PGY2s to also not show blanks
    ExtractTable.Range.AutoFilter Field:=PGY2ColNum, Criteria1 _
        :="<>"
        
    ExtractTable.Range.SpecialCells(xlCellTypeVisible).Copy
    
    DataWorkbook.Worksheets("PGY1&2").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' Unfilter original sheet
    Application.CutCopyMode = False

    
    ' Close Lookup Tables
    RotCoordLookupTable.Close
    EPALookupTable.Close
    
    
    '****************************************
    '   Delete Electives and Research Blocks
    '****************************************
    
    Dim rotationCol As Integer
    Dim lastRow As Integer
    Dim j As Integer
    Dim cellContents As String
    
    rotationCol = ExtractTable.ListColumns("Rotation").Range.Column
    
    ' PGY1 Sheet
    lastRow = DataWorkbook.Worksheets("PGY1").Cells(1, rotationCol).End(xlDown).Row
    
    For j = 2 To lastRow
        cellContents = UCase(DataWorkbook.Worksheets("PGY1").Cells(j, rotationCol).Value)
        If (InStr(cellContents, "ELECTIVE") > 0 Or InStr(cellContents, "RESEARCH") > 0) Then
            DataWorkbook.Worksheets("PGY1").Cells(j, rotationCol).EntireRow.Delete
            j = j - 1
            lastRow = lastRow - 1
        End If
    Next j
    
    ' PGY2 Sheet
    lastRow = DataWorkbook.Worksheets("PGY2").Cells(1, rotationCol).End(xlDown).Row
    
    For j = 2 To lastRow
        cellContents = UCase(DataWorkbook.Worksheets("PGY2").Cells(j, rotationCol).Value)
        If (InStr(cellContents, "ELECTIVE") > 0 Or InStr(cellContents, "RESEARCH") > 0) Then
            DataWorkbook.Worksheets("PGY2").Cells(j, rotationCol).EntireRow.Delete
            j = j - 1
            lastRow = lastRow - 1
        End If
    Next j
    
    ' PGY1&2 Sheet
    lastRow = DataWorkbook.Worksheets("PGY1&2").Cells(1, rotationCol).End(xlDown).Row
    
    For j = 2 To lastRow
        cellContents = UCase(DataWorkbook.Worksheets("PGY1&2").Cells(j, rotationCol).Value)
        If (InStr(cellContents, "ELECTIVE") > 0 Or InStr(cellContents, "RESEARCH") > 0) Then
            DataWorkbook.Worksheets("PGY1&2").Cells(j, rotationCol).EntireRow.Delete
            j = j - 1
            lastRow = lastRow - 1
        End If
    Next j
    
End Sub



