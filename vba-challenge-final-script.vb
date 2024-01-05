Sub ExtractDataByTicker()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerColumn As Range
    Dim uniqueTickers As Object
    Dim ticker As Variant
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        If Not IsNumeric(ws.Name) Then
            GoTo NextSheet
        End If
        
        ' Reset summary variables for each sheet
        Dim greatestIncreaseTicker As Variant
        Dim greatestDecreaseTicker As Variant
        Dim greatestVolumeTicker As Variant
        Dim greatestIncrease As Double
        Dim greatestDecrease As Double
        Dim greatestVolume As Double
        
        ' Find the last row in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Set ticker value range
        Set tickerColumn = ws.Range("A2:A" & lastRow)
        
        ' Initialize a Dictionary to store unique tickers
        Set uniqueTickers = CreateObject("Scripting.Dictionary")
        
        ' Loop through each cell in the ticker column
        For Each cell In tickerColumn
            If cell.Value <> "" And Not uniqueTickers.Exists(cell.Value) Then
                uniqueTickers.Add cell.Value, Nothing
            End If
        Next cell
        
        ' Output headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize the output row
        Dim outputRow As Long
        outputRow = 2
        
        ' Loop through each unique ticker
        For Each ticker In uniqueTickers.keys
            ' Call needed functions for each ticker
            ProcessTickerData ws, ticker, outputRow
            
            ' Check and update summary values
            Dim percentChange As Double
            percentChange = ws.Cells(outputRow, 11).Value
            
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            ElseIf percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            Dim totalVolume As Double
            totalVolume = ws.Cells(outputRow, 12).Value
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
            
            ' Move to the next row for output
            outputRow = outputRow + 1
        Next ticker
        
        ' Apply conditional formatting to the "Yearly Change" column
        ApplyConditionalFormatting ws
        
        ' Output summary values
        OutputSummary ws, greatestIncreaseTicker, greatestDecreaseTicker, greatestVolumeTicker, greatestIncrease, greatestDecrease, greatestVolume

        ' Reset outputRow for the next sheet
        outputRow = 2

NextSheet:
    Next ws
End Sub

Sub OutputSummary(ws As Worksheet, incTicker As Variant, decTicker As Variant, volTicker As Variant, incValue As Double, decValue As Double, volValue As Double)
    ' Initialize the output row for summary
    Dim summaryRow As Long
    summaryRow = 2
    
    ' Add headers for the summary section
    ws.Cells(1, 15).Value = "Summary"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ' Output summary values
    ws.Cells(summaryRow, 16).Value = incTicker
    ws.Cells(summaryRow, 17).Value = incValue
    summaryRow = summaryRow + 1
    
    ws.Cells(summaryRow, 16).Value = decTicker
    ws.Cells(summaryRow, 17).Value = decValue
    summaryRow = summaryRow + 1
    
    ws.Cells(summaryRow, 16).Value = volTicker
    ws.Cells(summaryRow, 17).Value = volValue
End Sub

Sub ProcessTickerData(ws As Worksheet, ticker As Variant, outputRow As Long)
    Dim firstRow As Long
    Dim lastRow As Long
    
    ' Find the first occurrence of the current ticker
    On Error Resume Next
    firstRow = WorksheetFunction.Match(ticker, ws.Columns("A"), 0)
    On Error GoTo 0
    
    ' Find the last occurrence of the current ticker
    On Error Resume Next
    lastRow = WorksheetFunction.Match(ticker, ws.Columns("A"), 1)
    On Error GoTo 0
    
    ' Check if both first and last rows are found
    If firstRow > 0 And lastRow > 0 Then
        firstRow = ws.Cells(firstRow, "A").Row
        lastRow = ws.Cells(lastRow, "A").Row
        
        ' Calculate yearly change, percent change, and total volume
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalVolume As Double
        
        yearlyChange = ws.Cells(lastRow, "F").Value - ws.Cells(firstRow, "C").Value
        If ws.Cells(firstRow, "C").Value <> 0 Then
            percentChange = yearlyChange / ws.Cells(firstRow, "C").Value
        Else
            percentChange = 0
        End If
        totalVolume = Application.WorksheetFunction.Sum(ws.Range("G" & firstRow & ":G" & lastRow))
        
        ' Output values in columns I, J, K, and L
        ws.Cells(outputRow, 9).Value = ticker
        ws.Cells(outputRow, 10).Value = yearlyChange
        ws.Cells(outputRow, 11).Value = percentChange
        ws.Cells(outputRow, 12).Value = totalVolume
    Else
        ' Handle possible case where the first or last row is not found
        Debug.Print "First or last row not found for ticker: " & ticker
    End If
End Sub

Sub ApplyConditionalFormatting(ws As Worksheet)
    Dim lastRow As Long
    
    ' Find the last row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Set the "Yearly Change" column range
    Dim rangeToFormat As Range
    Set rangeToFormat = ws.Range("J2:J" & lastRow)
    
    ' Apply conditional formatting for positive and negative values
    With rangeToFormat.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(0, 255, 0) ' Green for positive changes
    End With
    
    With rangeToFormat.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 0, 0) ' Red for negative changes
    End With
End Sub
