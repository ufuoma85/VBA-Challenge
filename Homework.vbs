Sub sortData()
Dim ws As Worksheet
Dim i As Long
Dim lastRow As Single
Dim totalTicker As Single
Dim openCellCount As Single
Dim closeCellCount As Single
Dim myRange As Range
Dim greatestIncreaseValue As Single
Dim greatestIncreaseTinker As String
Dim greatestDecreaseValue As Single
Dim greatestDecreaseTinker As String
Dim greatestTotalValue As Single
Dim greatestTotalTinker As String

'Application.ScreenUpdating = False
For Each Sheet In Worksheets
    Set ws = ThisWorkbook.Worksheets(Sheet.Name)
    lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    ws.Range("I1:Q100000").Clear
    
    ws.Range("A2:A" & lastRow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("I1"), Unique:=True 'Filter unique Ticker
    
    'Insert coloum head
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    
    totalTicker = ws.Range("I" & Rows.Count).End(xlUp).Row 'count total ticker
    openCellCount = 2 'data start from row 2
    
    For i = 2 To totalTicker
        ws.Range("M" & i) = Application.WorksheetFunction.CountIf(ws.Range("A2:A" & lastRow), ws.Range("I" & i)) 'Count each unique ticker number and insert in a supporting coloum
    
        closeCellCount = openCellCount - 1 + ws.Range("M" & i)
        'Write formula
        ws.Range("J" & i) = ws.Range("F" & closeCellCount) - ws.Range("C" & openCellCount) '=F252-C2
        ws.Range("K" & i) = Format(ws.Range("J" & i) / ws.Range("C" & openCellCount) * 100, "#.00") & "%" ' =J2/C2*100
        ws.Range("L" & i) = Application.WorksheetFunction.Sum(ws.Range("G" & openCellCount & ":G" & closeCellCount)) ' ='=SUM(G2:G252)
        openCellCount = closeCellCount + 1
    Next i
    
    'Search Greatest Increase Value
    Set myRange = ws.Range("K2:K" & totalTicker)
    greatestIncreaseValue = 0 'set "inital" greatestIncreaseValue
    For Each cell In myRange.Cells
        If cell.Value > greatestIncreaseValue Then 'if a value is larger than the old greatestIncreaseValue,
            greatestIncreaseValue = cell.Value ' store it as the new greatestIncreaseValue
            greatestIncreaseTinker = cell.Offset(0, -2).Value
        End If
    Next cell
    
    'Search Greatest Decrease Value
    Set myRange = ws.Range("K2:K" & totalTicker)
    greatestDecreaseValue = 1000
    For Each cell In myRange.Cells
        If cell.Value < greatestDecreaseValue Then
            greatestDecreaseValue = cell.Value
            greatestDecreaseTinker = cell.Offset(0, -2).Value
        End If
    Next cell
    
    'Search Greatest Total Volume
    Set myRange = ws.Range("L2:L" & totalTicker)
    greatestTotalValue = 0
    For Each cell In myRange.Cells
        If cell.Value > greatestTotalValue Then
            greatestTotalValue = cell.Value
            greatestTotalTinker = cell.Offset(0, -3).Value
        End If
    Next cell
           
    'Insert Result
    ws.Range("P2") = greatestIncreaseTinker
    ws.Range("Q2") = Format(greatestIncreaseValue * 100, "#.00") & "%"
    ws.Range("P3") = greatestDecreaseTinker
    ws.Range("Q3") = Format(greatestDecreaseValue * 100, "#.00") & "%"
    ws.Range("P4") = greatestTotalTinker
    ws.Range("Q4") = greatestTotalValue
    
    ws.Range("M1:M1000").Clear 'Clear supporting column data
    ws.Columns("I:Q").AutoFit
    
    'Conditional formating
    Set myRange = ws.Range("J2:J" & totalTicker)
    'Delete previous conditional formats
    myRange.FormatConditions.Delete
    myRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    myRange.FormatConditions(1).Interior.Color = vbRed
    myRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
    myRange.FormatConditions(2).Interior.Color = vbGreen

Next Sheet

End Sub







