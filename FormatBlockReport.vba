Option Explicit

' Excel won't let you directly call a macro that has a parameter,
' so this is a wrapper sub around Format_Block_Report so it will
' be listed in the Macro table
Sub FormatBlockReport()
    Application.ScreenUpdating = False
    Call Format_Block_Report
    Application.ScreenUpdating = True
End Sub

Sub Format_Block_Report(Optional extract As Workbook)

    Dim fso As Scripting.FileSystemObject
    Dim fd As FileDialog
    Dim fileWasChosen As Boolean
    Dim DataWorkbook As Workbook
    Dim DataSheet As Worksheet
    Dim RotCoordLookupTable As Workbook
    Dim EPALookupTable As Workbook
    Dim ExtractTable As ListObject
    Dim BlockReportCopy As String
    Dim firstYear As String
    Dim secondYear As String
    Dim newCol As Range
    Dim newHeader As Range
    Dim formulaString As String
    
    Set fso = New Scripting.FileSystemObject
    If extract Is Nothing Then
        MsgBox "Choose a Block Report to generate from..."
        Set fd = Application.FileDialog(msoFileDialogOpen)
    
        fd.Filters.Clear
        fd.Filters.Add "CSV Files", "*.csv, *.xl*"
        fd.FilterIndex = 1
    
        fd.AllowMultiSelect = False
        fd.InitialFileName = "N:\DOM1\Post Graduate Program\IM\Competence by Design - IM\Communications\Rotation Coordinator"
        fd.Title = "Choose an Extract Data File"
    
        fileWasChosen = fd.Show
    
        If Not fileWasChosen Then
            MsgBox "You didn't choose a Block Report File. Report generation was terminated."
            End
        End If
    
        Set DataWorkbook = Workbooks.Open(Filename:=fd.SelectedItems(1))
    Else
        Set DataWorkbook = extract
    End If
    
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
    
    'Select the RotCoordEmailData sheet as the Active Sheet (assumes sheet will be named this)
    DataWorkbook.Worksheets("RotCoordEmailData").Activate
    
    ' Rename the workbook copy with the appropriate division name and date range
    DataWorkbook.SaveAs DataWorkbook.Path & "\" & "Block " & DataWorkbook.ActiveSheet.Range("A3").Value & _
        " Rotation Coordinator " & firstYear & "-" & secondYear & ".xlsx", 51
    Set DataSheet = DataWorkbook.Worksheets("RotCoordEmailData")
    DataSheet.Name = "OriginalSheet"
    
    ' Create a table from the data for ease of manipulation - allows use of column
    ' header names for flexibility
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes).Name = _
        "ExtractTable"
    Set ExtractTable = ActiveSheet.ListObjects("ExtractTable")
    
    '****************************************
    '   Delete Electives, Research Blocks, and Leaves
    '****************************************
    
    Dim rotationCol As Integer
    Dim lastRow As Integer
    Dim j As Integer
    Dim cellContents As String
    
    rotationCol = ExtractTable.ListColumns("Rotation").Range.Column
    
    lastRow = DataWorkbook.Worksheets("OriginalSheet").Cells(1, rotationCol).End(xlDown).Row
    
    For j = 2 To lastRow
        cellContents = UCase(DataWorkbook.Worksheets("OriginalSheet").Cells(j, rotationCol).Value)
        If (InStr(cellContents, "ELECTIVE") > 0 Or InStr(cellContents, "RESEARCH") > 0 Or InStr(cellContents, "LEAVE") > 0) Then
            DataWorkbook.Worksheets("OriginalSheet").Cells(j, rotationCol).EntireRow.Delete
            j = j - 1
            lastRow = lastRow - 1
        End If
    Next j
    
    ' Replace GIM-CTU/Consults with GIM-CTU - Consults because / is replaced in Lookup table ( "/" is invalid in file names derived from rotation)
    DataSheet.Range(Cells(1, rotationCol), Cells(lastRow, rotationCol)).Select
    Selection.Replace What:="GIM - CTU/Consults Experience", Replacement:="GIM - CTU - Consults Experience", LookAt:=xlWhole, MatchCase:=False
    Selection.Replace What:="GIM - CTU/Junior Experience", Replacement:="GIM - CTU - Junior Experience", LookAt:=xlWhole, MatchCase:=False
    
    

    ' Insert Column for creating the Junior Rotation-Stage EPA lookup key
    ExtractTable.ListColumns("PGY1s").Range.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "RotationStageJunior"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=[@Rotation]&IF([@Period] < 5, " & Chr(34) & "TTD" & Chr(34) & ", " & Chr(34) & "FOD" & Chr(34) & ")"
    ExtractTable.ListColumns("RotationStageJunior").Range.EntireColumn.AutoFit
    
    ' Insert Column for creating the Senior Rotation-Stage EPA lookup key
    ExtractTable.ListColumns("PGY1s").Range.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "RotationStageSenior"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=[@Rotation]&" & Chr(34) & "COD" & Chr(34)
    ExtractTable.ListColumns("RotationStageSenior").Range.EntireColumn.AutoFit
    
    
    ' Insert Column to denote rotations with both Juniors and Seniors
    DataWorkbook.Worksheets("OriginalSheet").Activate
    ExtractTable.ListColumns("PGY1s").Range.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "JuniorAndSeniorRotation"
    ActiveCell.Offset(1, 0).Select
    'Next line Formula is =AND([@PGY1s<>"", OR([@PGY2s<>"", [@PGY3s]<>""))
    ActiveCell.FormulaR1C1 = "=AND([@PGY1s]<>" & Chr(34) & Chr(34) & ", OR([@PGY2s]<>" & Chr(34) & Chr(34) & ", [@PGY3s]<>" & Chr(34) & Chr(34) & "))"
    ExtractTable.ListColumns("JuniorAndSeniorRotation").Range.EntireColumn.AutoFit
    
    
    ExtractTable.Range.Cells.Select
    ' Replace all NULLS
    Selection.Replace What:="NULL", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False
    ' Break names onto new line by comma-space
    Selection.Replace What:=", ", Replacement:="" & Chr(10) & "", LookAt:=xlPart, _
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
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY1 Priority (Highest*)"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@RotationStageJunior],'[" & EPALookupTable.Name & "]Sheet1'!$C:$G,2,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY1 Always do when you can"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@RotationStageJunior],'[" & EPALookupTable.Name & "]Sheet1'!$C:$G,3,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY1 Optional"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@RotationStageJunior],'[" & EPALookupTable.Name & "]Sheet1'!$C:$G,4,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY2-3 Priority (Highest*)"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@RotationStageSenior],'[" & EPALookupTable.Name & "]Sheet1'!$C:$G,2,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY2-3 Always do when you can"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@RotationStageSenior],'[" & EPALookupTable.Name & "]Sheet1'!$C:$G,3,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY2-3 Optional"
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@RotationStageSenior],'[" & EPALookupTable.Name & "]Sheet1'!$C:$G,4,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY1 Rotation Cards "
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@RotationStageJunior],'[" & EPALookupTable.Name & "]Sheet1'!$C:$G,5,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(0, 1).FormulaR1C1 = "PGY2-3 Rotation Cards "
    ' Note: Chr(34) is a function returning the quote character
    formulaString = "=IFERROR(VLOOKUP([@RotationStageSenior],'[" & EPALookupTable.Name & "]Sheet1'!$C:$G,5,FALSE)," & Chr(34) & Chr(34) & ")"
    ExtractTable.HeaderRowRange.End(xlToRight).Offset(1, 0).Formula = formulaString
    
    ' Copy and paste all values so cells don't contain formulas
    ExtractTable.Range.SpecialCells(xlCellTypeVisible).Copy
    
    DataWorkbook.Worksheets("OriginalSheet").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    ' Clear out any 0s that were returned by the VLOOKUPs
    DataWorkbook.Worksheets("OriginalSheet").Cells.Replace What:="0", Replacement:="", _
        LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, _
        SearchFormat:=False, ReplaceFormat:=False
    
    
    '****************************
    '   Copy PGY1s to New Sheet
    '****************************
    Dim PGY1ColNum As Integer
    Dim PGY2ColNum As Integer
    Dim PGY3ColNum As Integer
    Dim JuniorSeniorRotationColNum As Integer
    
    PGY1ColNum = ExtractTable.ListColumns("PGY1s").Range.Column
    PGY2ColNum = ExtractTable.ListColumns("PGY2s").Range.Column
    PGY3ColNum = ExtractTable.ListColumns("PGY3s").Range.Column
    JuniorSeniorRotationColNum = ExtractTable.ListColumns("JuniorAndSeniorRotation").Range.Column
    
    DataWorkbook.Worksheets.Add After:=DataWorkbook.Worksheets("OriginalSheet")
    If extract Is Nothing Then
        DataWorkbook.Worksheets("Sheet1").Name = "PGY1"
    Else
        DataWorkbook.Worksheets("Sheet3").Name = "PGY1"
    End If
    
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
    '   Copy PGY2-3s to New Sheet
    '****************************
    
    DataWorkbook.Worksheets.Add After:=DataWorkbook.Worksheets("PGY1")
    If extract Is Nothing Then
        DataWorkbook.Worksheets("Sheet2").Name = "PGY2-3"
    Else
        DataWorkbook.Worksheets("Sheet4").Name = "PGY2-3"
    End If
    
    ' Filter PGY1s to only show blanks
    ExtractTable.Range.AutoFilter Field:=PGY1ColNum, Criteria1 _
        :="="
        
    ExtractTable.Range.SpecialCells(xlCellTypeVisible).Copy
    
    DataWorkbook.Worksheets("PGY2-3").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    '****************************
    '   Copy PGY1&2-3s to New Sheet
    '****************************
    
    
    DataWorkbook.Worksheets.Add After:=DataWorkbook.Worksheets("PGY2-3")
    If extract Is Nothing Then
        DataWorkbook.Worksheets("Sheet3").Name = "PGY1&2-3"
    Else
        DataWorkbook.Worksheets("Sheet5").Name = "PGY1&2-3"
    End If
    
    
    ' Filter ExtractTable to rotations with juniors and seniors
    ExtractTable.AutoFilter.ShowAllData
    ExtractTable.Range.AutoFilter Field:=JuniorSeniorRotationColNum, Criteria1 _
        :="TRUE"
        
    ExtractTable.Range.SpecialCells(xlCellTypeVisible).Copy
    
    DataWorkbook.Worksheets("PGY1&2-3").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' Unfilter original sheet
    Application.CutCopyMode = False
    DataSheet.Activate
    ExtractTable.AutoFilter.ShowAllData
    
    ' Close Lookup Tables
    RotCoordLookupTable.Close
    EPALookupTable.Close
    
    Range("A1").Select
    Application.CutCopyMode = False
    
    ' If the extract is passed as a parameter, and thus this sub is being called from
    ' GenerateRotCoordReport, then save the formatted extract so it can be inspected
    ' afterwards if needed
    
    If Not extract Is Nothing Then
        DataWorkbook.Save
    End If
End Sub



