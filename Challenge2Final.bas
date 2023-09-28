Attribute VB_Name = "Module1"
Sub Challenge2():

For Each ws In Worksheets
    Dim sheetName As String
    sheetName = ws.Name


    ' add column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    

    ' initialize variables
    Dim tickerName As String
    Dim yearlyChange As Double   ' yearly change dollar amount
        yearlyChange = 0
    Dim percentChange As Double  ' yearly change percent
        percentChange = 0
    Dim totalVolume As LongLong
        totalVolume = 0
    Dim tickerCount As Integer
        tickerCount = 2
    Dim yearOpen As Double      ' first open price of the year
    Dim yearClose As Double     ' last close price of the year
    
    ' find last row
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    ' create summary rows with ticker name and total volume
    For Col = 1 To 1
    tickerName = ws.Cells(2, 1).Value
    yearOpen = ws.Cells(2, 3).Value    ' finding yearOpen and yearClose works because the entries are in chronological order
    
        For Row = 2 To lastRow
        
            If ws.Cells(Row + 1, Col).Value <> ws.Cells(Row, Col).Value Then
                ' add ticker name to summary column
                ws.Cells(tickerCount, 9).Value = tickerName
                
                ' add total volume to summary
                ws.Cells(tickerCount, 12).Value = totalVolume
                
                ' get this year's close
                yearClose = ws.Cells(Row, 6).Value
                ' calculate yearly change
                yearlyChange = yearClose - yearOpen
                ' put it in the summary column
                ws.Cells(tickerCount, 10).Value = yearlyChange
                
                ' conditional formatting for year change
                    If ws.Cells(tickerCount, 10).Value < 0 Then
                        ws.Cells(tickerCount, 10).Interior.ColorIndex = 3
                    ElseIf ws.Cells(tickerCount, 10).Value > 0 Then
                        ws.Cells(tickerCount, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(tickerCount, 10).Interior.ColorIndex = 27
                    End If
                        
                
                ' calculate percent change
                percentChange = yearClose / yearOpen
                ' put it in the summary column
                ws.Cells(tickerCount, 11).Value = percentChange - 1
                    
                
                
                
                ' go to next row
                tickerCount = tickerCount + 1
                
                ' reset total volume
                totalVolume = 0
                
                ' set new ticker name
                tickerName = ws.Cells(Row + 1, Col).Value
                
                ' get next ticker's year open
                yearOpen = ws.Cells(Row + 1, 3).Value
            Else
                totalVolume = totalVolume + ws.Cells(Row, Col + 6).Value
            End If
        
        Next Row
    Next Col
    
    
    
    ' greatest % summary columns
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    
    ' find max % inc, % dec, total vol
    Dim maxPercentInc As Double
    Dim maxPercentDec As Double
    Dim maxTotalVol As LongLong
    Dim maxPercentIncIndex, maxPercentDecIndex, maxTotalVolIndex As Integer
    
    ' fill in Value column in second summary column
    maxPercentInc = WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(2, 17).Value = maxPercentInc
    
    maxPercentDec = WorksheetFunction.Min(ws.Range("K:K"))
    ws.Cells(3, 17).Value = maxPercentDec
    
    maxTotalVol = WorksheetFunction.Max(ws.Range("L:L"))
    ws.Cells(4, 17).Value = maxTotalVol
    
    
    ' fill in Ticker column in second summary column
    maxPercentIncIndex = WorksheetFunction.Match(maxPercentInc, ws.Range("K:K"), 0)
    ws.Cells(2, 16).Value = ws.Range("I" & maxPercentIncIndex).Value
    
    maxPercentDecIndex = WorksheetFunction.Match(maxPercentDec, ws.Range("K:K"), 0)
    ws.Cells(3, 16).Value = ws.Range("I" & maxPercentDecIndex).Value
    
    maxTotalVolIndex = WorksheetFunction.Match(maxTotalVol, ws.Range("L:L"), 0)
    ws.Cells(4, 16).Value = ws.Range("I" & maxTotalVolIndex).Value
    
    
    
    
    
    



    ' formatting
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("K:K").NumberFormat = "0.00%"
    ws.Range("I1:L1").Columns.AutoFit ' this should be last
    ws.Range("O1:Q4").Columns.AutoFit
    
Next ws
End Sub

