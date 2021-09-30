Attribute VB_Name = "Module1"
Sub createSummaryTableOnEverySheet()
'Note to self: this method assumes that data is 1) sorted by ticker, 2) then sorted by date/year, 3) starting and closing prices cannot be negative

    'header storage variables
        'main assignment
            Dim mainHeaders(3) As String
            mainHeaders(0) = "Ticker"
            mainHeaders(1) = "Yearly Change"
            mainHeaders(2) = "Percent Change"
            mainHeaders(3) = "Total Stock Volume"
        'bonus assignment
            Dim bonusHeaders(4) As String
            bonusHeaders(0) = "Greatest % Increase"
            bonusHeaders(1) = "Greatest % Decrease"
            bonusHeaders(2) = "Greatest Total Volume"
            bonusHeaders(3) = "Ticker"
            bonusHeaders(4) = "Value"
            
    'worksheet looping variables
        Dim wb As Workbook
        Set wb = ThisWorkbook
        Dim ws As Worksheet
    
    'create headers on every sheet
        'creating main assignment headers
        For Each ws In wb.Worksheets
            Dim i As Integer
            For i = 0 To 3
                ws.Cells(1, i + 9).Value = mainHeaders(i)
                ws.Cells(1, i + 9).ColumnWidth = 16
                ws.Cells(1, i + 9).Interior.ColorIndex = 1
                ws.Cells(1, i + 9).Font.ColorIndex = 2
            Next i
        Next ws
        
        'creating bonus assignment headers
        For Each ws In wb.Worksheets
            Dim y As Integer
            For y = 0 To 4
                If (y < 3) Then
                    ws.Cells(y + 2, 14).Value = bonusHeaders(y)
                    ws.Cells(y + 2, 14).ColumnWidth = 19
                    ws.Cells(y + 2, 14).Interior.ColorIndex = 1
                    ws.Cells(y + 2, 14).Font.ColorIndex = 2
                Else
                    ws.Cells(1, y + 12).Value = bonusHeaders(y)
                    ws.Cells(1, y + 12).ColumnWidth = 16
                    ws.Cells(1, y + 12).Interior.ColorIndex = 1
                    ws.Cells(1, y + 12).Font.ColorIndex = 2
                End If
            Next y
        Next ws
    
    'populate main assignment headers
        'storage variables for calculations done during populating
        Dim openPrice, closePrice As Double
        Dim sumStock As LongLong
        Dim nextEmptySummaryRow As Long
        
        'looping variables
        Dim row, lastRow As Long
        
        'populating yearly, % change and total stock
        For Each ws In wb.Worksheets
            'set up
            nextEmptySummaryRow = 2 'reset this row for every sheet
            openPrice = ws.Cells(2, 3).Value 'set opening price for first ticker for every sheet
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row 'recalculate last row for every sheet
            
            'loop
            For row = 2 To lastRow
                If (ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value) Then 'if found a unique ticker...
                    'mark down its closing price and push to summary table
                    closePrice = ws.Cells(row, 6).Value
                    ws.Cells(nextEmptySummaryRow, 9).Value = ws.Cells(row, 1).Value
                    
                    'conditionally format its yearly price and then push to summary table
                    If (closePrice - openPrice > 0) Then
                        ws.Cells(nextEmptySummaryRow, 10).Interior.ColorIndex = 4
                    ElseIf (closePrice - openPrice < 0) Then
                        ws.Cells(nextEmptySummaryRow, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(nextEmptySummaryRow, 10).Interior.ColorIndex = 15
                    End If
                    ws.Cells(nextEmptySummaryRow, 10).Value = closePrice - openPrice
                    
                    'check for special cases of percent change, then push percent change
                    If (openPrice <> 0) Then 'if openPrice != 0
                        ws.Cells(nextEmptySummaryRow, 11).Value = FormatPercent((closePrice - openPrice) / openPrice)
                    ElseIf (closePrice = 0) Then 'if openPrice = 0 AND closePrice = 0
                        ws.Cells(nextEmptySummaryRow, 11).Value = 0
                    Else 'if only openPrice = 0
                        ws.Cells(nextEmptySummaryRow, 11).Value = "Infinite (openPrice started as 0)"
                    End If
                    
                    'push ticker's total stock; this has been calculated by the time a unique ticker is found; code in the ELSE part of this if statement (see below)
                    ws.Cells(nextEmptySummaryRow, 12).Value = sumStock
                    
                    'clean up
                    openPrice = ws.Cells(row + 1, 3).Value 'set openPrice to the next ticker's opening price, calculation used in next iteration of this loop
                    nextEmptySummaryRow = nextEmptySummaryRow + 1 'increment so pushed values dont overwrite existing ones
                    sumStock = 0 'reset sumStock for the next unique ticker
                Else 'if didn't find unique ticker...
                    'sum up all total stocks, since we're still in the range of the same ticker
                    sumStock = sumStock + ws.Cells(row, 7).Value
                End If
            Next row
        Next ws
        
    'populate bonus assignment headers
        'storage variables
        Dim keyValues(5) As Variant
        Dim keyValuesIndex As Integer
    
        'populating each summary table on each worksheet
        For Each ws In wb.Worksheets
            'initializing the starter values for every sheet
                '(0) = highestPercentIncreaseTicker (string)
                '(1) = highestPercentDecreaseTicker (string)
                '(2) = highestVolumeTicker (string)
                '(3) = highestPercentIncrease (double)
                '(4) = highestPercentDecrease (double)
                '(5) = highestVolume (longlong)
                For keyValuesIndex = 0 To 5
                    If (keyValuesIndex < 3) Then 'when keyValuesIndex = 0 to 2
                        keyValues(keyValuesIndex) = ws.Cells(2, 9).Value 'set to ticker "A"
                    ElseIf (keyValuesIndex <> 5) Then 'when keyValuesIndex = 3 to 4
                        keyValues(keyValuesIndex) = ws.Cells(2, 11).Value 'set to ticker "A"'s percent change
                    Else 'when keyValuesIndex = 5
                        keyValues(keyValuesIndex) = ws.Cells(2, 12).Value 'set to ticker "A"'s total stock
                    End If
                Next keyValuesIndex
            
            'comparing rest of values to the starter values
                'Picking out key tickers
                    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
                    For row = 2 To lastRow
                        'HIGHEST PERCENT: if found a higher value than one in array, populate and markdown its ticker
                        If (ws.Cells(row, 11).Value > keyValues(3)) Then
                            keyValues(3) = ws.Cells(row, 11).Value
                            keyValues(0) = ws.Cells(row, 9).Value
                        End If
                        
                        'LOWEST PERCENT: if found a lower value than one in array, populate and markdown its ticker
                        If (ws.Cells(row, 11).Value < keyValues(4)) Then
                            keyValues(4) = ws.Cells(row, 11).Value
                            keyValues(1) = ws.Cells(row, 9).Value
                        End If
                        
                        'VOLUME: if found a higher value than one in array, populate and markdown its ticker
                        If (ws.Cells(row, 12).Value > keyValues(5)) Then
                            keyValues(5) = ws.Cells(row, 12).Value
                            keyValues(2) = ws.Cells(row, 9).Value
                        End If
                    Next row
                    
                'Populating key table
                    keyValuesIndex = 0
                    For Column = 15 To 16
                        For row = 2 To 4
                            ws.Cells(row, Column).Value = keyValues(keyValuesIndex)
                            keyValuesIndex = keyValuesIndex + 1
                        Next row
                    Next Column
        Next ws
End Sub

