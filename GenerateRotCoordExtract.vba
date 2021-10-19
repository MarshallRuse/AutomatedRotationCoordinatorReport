Option Explicit

' Excel won't let you directly call a macro that has a parameter,
' so this is a wrapper sub around Generate_Rotation_Coordinator_Extract so it will
' be listed in the Macro table
Sub GenerateRotationCoordinatorExtract()
    Application.ScreenUpdating = False
    Call Generate_Rotation_Coordinator_Extract
    Application.ScreenUpdating = True
End Sub



Function JoinByRotationLocationAndLevel( _
    Delimiter As Variant, _
    TextRange As Range, _
    rotationLocationLookupRange As Range, _
    RotationLocationComparator As String, _
    trainingLevelRange As Range, _
    TrainingLevel As Integer) As Variant
    
    Dim textarray() As Variant
    
    Dim i As Integer
    Dim k As Integer
    k = 0
    Dim l As Integer
    l = 0
    Dim residentString As String
    
    For i = 2 To TextRange.Cells.Count
        If TextRange.Cells(i) <> "" And _
            rotationLocationLookupRange.Cells(i) = RotationLocationComparator And _
            trainingLevelRange.Cells(i) = TrainingLevel Then
                k = k + 1
                ReDim Preserve textarray(1 To k)
                textarray(k) = TextRange.Cells(i)
        End If
    Next i
    
    'Now Join the Cells
    If Not Not textarray Then
        If Not TypeName(Delimiter) = "Range" Then
            JoinByRotationLocationAndLevel = textarray(1)
                For i = 2 To UBound(textarray) - 1
                JoinByRotationLocationAndLevel = JoinByRotationLocationAndLevel & Delimiter & textarray(i)
                Next i
            If i > 1 Then JoinByRotationLocationAndLevel = JoinByRotationLocationAndLevel & Delimiter & textarray(UBound(textarray))
        Else
           JoinByRotationLocationAndLevel = textarray(1)
                For i = 2 To UBound(textarray) - 1
                    l = l + 1
                    If l = Delimiter.Cells.Count + 1 Then l = 1
                JoinByRotationLocationAndLevel = JoinByRotationLocationAndLevel & Delimiter.Cells(l) & textarray(i)
                Next i
            If i > 1 Then JoinByRotationLocationAndLevel = JoinByRotationLocationAndLevel & Delimiter.Cells(l + i) & textarray(UBound(textarray))
        End If
    Else
        JoinByRotationLocationAndLevel = ""
    End If
End Function



Sub Generate_Rotation_Coordinator_Extract(Optional ByRef extract As Workbook)

    Dim fso As Scripting.FileSystemObject
    Dim fd As FileDialog
    Dim fileWasChosen As Boolean
    Dim DataWorkbook As Workbook
    Dim DataSheet As Worksheet
    Dim ExtractTable As ListObject
    Dim firstYear As String
    Dim secondYear As String
    Dim newCol As Range
    Dim newHeader As Range
    Dim finalSheet As Worksheet

    Application.ScreenUpdating = False
    
    Set fso = New Scripting.FileSystemObject
    
    If extract Is Nothing Then
        MsgBox "Choose an IRIS/ORBS extract to generate from..."
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
        Set extract = DataWorkbook
    
    Else
        Set DataWorkbook = extract
    End If
    
    ' Ask the User which Block this is for
    Dim UserValue As Variant
    Dim UserReply As Variant
    Dim BlockNumber As Integer
    Dim BlockNumberChosen As Boolean
    Dim BlockSelectionFinished As Boolean
    Do
        BlockNumberChosen = False
        BlockSelectionFinished = False
        
        UserReply = MsgBox(Prompt:="Which Block is this report for? (ex. '1', '2', '12')", Buttons:=vbOKCancel, Title:="Choose Block")
        If UserReply = vbOK Then
    
            UserValue = InputBox(Prompt:="Enter an integer Block number, otherwise click Cancel", _
                                             Title:="Block Number")
            If UserValue = vbNullString Then
                BlockNumberChosen = False
                BlockSelectionFinished = True
                ' do nothing
            ElseIf Not IsNumeric(UserValue) Then
                    MsgBox "You must enter a numeric value"
            ElseIf Len(UserValue) > 2 Then
                    MsgBox "Max Block number is 2 digits", vbCritical
            Else
                BlockNumber = CInt(UserValue)
                If BlockNumber > 13 Or BlockNumber < 1 Then
                    MsgBox "Blocks must be between 1 and 13, inclusive.", vbCritical
                Else
                    BlockNumberChosen = True
                    BlockSelectionFinished = True
                End If
            End If
        Else
            BlockSelectionFinished = True
        End If
        
    Loop Until BlockNumberChosen = True Or BlockSelectionFinished = True
    
    If Not BlockNumberChosen Then
        MsgBox "You didn't choose a Block to generate an extract from. Rotation Coordinator Extract generation was terminated."
        End
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
    
    'Select the Sheet1 sheet as the Active Sheet (assumes sheet will be named this)
    DataWorkbook.Worksheets("Sheet1").Activate
    
    ' Rename the workbook copy with the appropriate division name and date range
    DataWorkbook.SaveAs DataWorkbook.Path & "\" & "Rotation Coordinator Data - Block " & BlockNumber & _
        " " & firstYear & "-" & secondYear & " - " & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & ".xlsx", 51
    Set DataSheet = DataWorkbook.Worksheets("Sheet1")
    DataSheet.Name = "ExtractData"
    
    ' Create a table from the data for ease of manipulation - allows use of column
    ' header names for flexibility
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1").CurrentRegion, , xlYes).Name = _
        "ExtractTable"
    Set ExtractTable = ActiveSheet.ListObjects("ExtractTable")
    
    ' Copy the Block lookup table into the DataWorkbook
    DataWorkbook.Worksheets.Add After:=DataWorkbook.Worksheets("ExtractData")
    DataWorkbook.Worksheets("Sheet1").Name = "BLOCK"
    
    ThisWorkbook.Worksheets("BLOCK").Range("A1").CurrentRegion.Copy DataWorkbook.Worksheets("BLOCK").Range("A1")
    
    ' Use the BLOCK lookup table to lookup the rotation start date, name this column Period
    DataWorkbook.Worksheets("ExtractData").Activate
    ExtractTable.ListColumns("Rotation").Range.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "Period"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Formula = "=VLOOKUP([@RotationStartDate], 'BLOCK'!A:C, 3, TRUE)"
    ExtractTable.ListColumns("Period").Range.EntireColumn.AutoFit
    
    ' Filter for everything BUT the desired block, delete all rows, unfilter to leave just desired block
    Dim PeriodColumn As Integer
    Dim i As Integer
    Dim lastRow As Integer
    Dim cellContents As Integer
    ReDim blocksToDelete(1 To 12) As String
    
    PeriodColumn = ExtractTable.ListColumns("Period").Range.Column
    lastRow = DataWorkbook.Worksheets("ExtractData").Cells(1, PeriodColumn).End(xlDown).Row
    
    For i = 1 To 13
        If Not i = BlockNumber Then
            If i > BlockNumber Then
                blocksToDelete(i - 1) = Str(i)
            Else
                blocksToDelete(i) = Str(i)
            End If
        End If
    Next i
    
    ExtractTable.Range.AutoFilter Field:=PeriodColumn, Criteria1 _
        :=blocksToDelete, Operator:= _
        xlFilterValues
    
    On Error Resume Next
        Application.DisplayAlerts = False
        ExtractTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Delete
        Application.DisplayAlerts = True
    On Error GoTo 0
    
    
    ExtractTable.Range.AutoFilter Field:=BlockNumber
    
    ' Add a Resident column
    ExtractTable.ListColumns("CurrentEmail").Range.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "Resident"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Formula = "=TRIM( PROPER( [@TraineeName] )& " & Chr(34) & " " & Chr(34) & " & UPPER( [@LastName] )& IF( [@Team] <> " & Chr(34) & Chr(34) & ", " & Chr(34) & " (" & Chr(34) & "& [@Team] & " & Chr(34) & ")" & Chr(34) & ", " & Chr(34) & Chr(34) & "))"
    ExtractTable.ListColumns("Resident").Range.EntireColumn.AutoFit
    
    
    ' Delete duplicate resident listings
    Dim ResidentColumn As Integer
    ResidentColumn = ExtractTable.ListColumns("Resident").Range.Column
    ActiveSheet.Range("ExtractTable[#All]").RemoveDuplicates Columns:=ResidentColumn, Header _
        :=xlYes
    
    ' Insert RotationLocationLookup column
    ExtractTable.ListColumns(1).Range.Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Set newCol = Selection.EntireColumn
    Set newHeader = newCol.Cells(1)
    newHeader.FormulaR1C1 = "RotationLocationLookup"
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.ListObjects("ExtractTable").Resize Range("$A$1:$K$217")
    Range("A2").Select
    ActiveCell.Formula = "=[@Rotation]&[@Location]"
    ExtractTable.ListColumns("RotationLocationLookup").Range.EntireColumn.AutoFit
    
    ' Add a Sheet named RotCoordEmailData
    DataWorkbook.Worksheets.Add Before:=DataWorkbook.Worksheets("ExtractData")
    DataWorkbook.Worksheets("Sheet2").Name = "RotCoordEmailData"
    Set finalSheet = DataWorkbook.Worksheets("RotCoordEmailData")
    
    ' Copy Period, Rotation, Location into columns A,B,C, respectively
    DataWorkbook.Worksheets("ExtractData").Activate
    ExtractTable.ListColumns("Period").Range.Copy
    finalSheet.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    DataWorkbook.Worksheets("ExtractData").Activate
    ExtractTable.ListColumns("Rotation").Range.Copy
    finalSheet.Range("B1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    DataWorkbook.Worksheets("ExtractData").Activate
    ExtractTable.ListColumns("Location").Range.Copy
    finalSheet.Range("C1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ' Rename Location to Hospital
    finalSheet.Range("C1").Value = "Hospital"
    
    ' Add Columns for PGY1-3s
    finalSheet.Range("D1").Value = "PGY1s"
    finalSheet.Range("E1").Value = "PGY1Emails"
    finalSheet.Range("F1").Value = "PGY2s"
    finalSheet.Range("G1").Value = "PGY2Emails"
    finalSheet.Range("H1").Value = "PGY3s"
    finalSheet.Range("I1").Value = "PGY3Emails"
    
    ' Remove duplicate rotation/locations
    finalSheet.Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(2, 3), _
        Header:=xlYes
        
    
    Dim rotationLocationLookupRange As Range
    Dim trainingLevelRange As Range
    Dim residentRange As Range
    Dim emailRange As Range
    Dim lastRowIndex As Integer
    Dim rotLoc As String
    
    DataWorkbook.Worksheets("ExtractData").Activate
    
    Set rotationLocationLookupRange = ExtractTable.ListColumns("RotationLocationLookup").Range
    Set trainingLevelRange = ExtractTable.ListColumns("TrainingLevel").Range
    Set residentRange = ExtractTable.ListColumns("Resident").Range
    Set emailRange = ExtractTable.ListColumns("CurrentEmail").Range
    
    DataWorkbook.Worksheets("RotCoordEmailData").Activate
    lastRowIndex = finalSheet.Range("A1").End(xlDown).Row
    
    Dim j As Integer
    ' PGY1s
    For j = 2 To lastRowIndex
        rotLoc = finalSheet.Cells(j, 2).Value & finalSheet.Cells(j, 3).Value
        finalSheet.Cells(j, 4).Value = JoinByRotationLocationAndLevel(", ", residentRange, rotationLocationLookupRange, rotLoc, trainingLevelRange, 1)
    Next j
    
    ' PGY1Emails
    For j = 2 To lastRowIndex
        rotLoc = finalSheet.Cells(j, 2).Value & finalSheet.Cells(j, 3).Value
        finalSheet.Cells(j, 5).Value = JoinByRotationLocationAndLevel(", ", emailRange, rotationLocationLookupRange, rotLoc, trainingLevelRange, 1)
    Next j
    
    'PGY2s
    For j = 2 To lastRowIndex
        rotLoc = finalSheet.Cells(j, 2).Value & finalSheet.Cells(j, 3).Value
        finalSheet.Cells(j, 6).Value = JoinByRotationLocationAndLevel(", ", residentRange, rotationLocationLookupRange, rotLoc, trainingLevelRange, 2)
    Next j
    
    ' PGY2Emails
    For j = 2 To lastRowIndex
        rotLoc = finalSheet.Cells(j, 2).Value & finalSheet.Cells(j, 3).Value
        finalSheet.Cells(j, 7).Value = JoinByRotationLocationAndLevel(", ", emailRange, rotationLocationLookupRange, rotLoc, trainingLevelRange, 2)
    Next j
    
    'PGY3s
    For j = 2 To lastRowIndex
        rotLoc = finalSheet.Cells(j, 2).Value & finalSheet.Cells(j, 3).Value
        finalSheet.Cells(j, 8).Value = JoinByRotationLocationAndLevel(", ", residentRange, rotationLocationLookupRange, rotLoc, trainingLevelRange, 3)
    Next j
    
    ' PGY3Emails
    For j = 2 To lastRowIndex
        rotLoc = finalSheet.Cells(j, 2).Value & finalSheet.Cells(j, 3).Value
        finalSheet.Cells(j, 9).Value = JoinByRotationLocationAndLevel(", ", emailRange, rotationLocationLookupRange, rotLoc, trainingLevelRange, 3)
    Next j
    
    Application.ScreenUpdating = True
End Sub
