Sub StockSummary()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
    
    ' Looping through each of the sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Finding the last row of data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Creating the table headers for requested values
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initializing variables
        summaryRow = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        tickerGreatestIncrease = ""
        tickerGreatestDecrease = ""
        tickerGreatestVolume = ""
        
        ' Looping through each row of data
        For i = 2 To lastRow
            ' Checking if the current ticker is the same as the previous ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Or i = 2 Then
                ' Storing the ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Storing the opening price
                openingPrice = ws.Cells(i, 3).Value
                
                ' Resetting the total volume
                totalVolume = 0
            End If
            
            ' Storing the closing price
            closingPrice = ws.Cells(i, 6).Value
            
            ' Calculating the yearly change
            yearlyChange = closingPrice - openingPrice
            
            ' Calculating the percent change
            If openingPrice <> 0 Then
                percentChange = yearlyChange / openingPrice * 100
            Else
                percentChange = 0
            End If
            
            ' Adding to total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Outputting results
            If ws.Cells(i + 1, 1).Value <> ticker Then
                
                ' Outputting results in summary table
                ws.Cells(summaryRow, 9).Value = ticker
                'ws.Cells(summaryRow, 10).Value = Format(yearlyChange, "0.00")
                ws.Cells(summaryRow, 10).NumberFormat = "0.00"
                ws.Cells(summaryRow, 10).Value = yearlyChange
                
                ws.Cells(summaryRow, 11).NumberFormat = "0.00\%"
                ws.Cells(summaryRow, 11).Value = percentChange
                
                ws.Cells(summaryRow, 12).NumberFormat = "0"
                ws.Cells(summaryRow, 12).Value = totalVolume
              
                ' Coloring cells based on yearly change. Red for decrease, green for increase
                If yearlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Checking for the greatest increase, decrease, and volume
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    tickerGreatestIncrease = ticker
                End If
                
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    tickerGreatestDecrease = ticker
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    tickerGreatestVolume = ticker
                End If
                
                ' Moving to the next row in the summary table
                summaryRow = summaryRow + 1
            End If
        Next i
        
        ' Outputting greatest increase, decrease, and volume
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = tickerGreatestIncrease
        ws.Cells(2, 17).Value = greatestIncrease & "%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = tickerGreatestDecrease
        ws.Cells(3, 17).Value = greatestDecrease & "%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = tickerGreatestVolume
        ws.Cells(4, 17).Value = greatestVolume
        ws.Cells(4, 17).NumberFormat = "0.00E+00"
        
        ' Auto Fitting columns
        ws.Columns("I:Q").AutoFit
    Next ws
    
End Sub
