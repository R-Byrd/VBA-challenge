Sub StockSummary()
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim ticker As String
    Dim openingrice As Double
    Dim closingrice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalvolume As Double
    Dim summaryrow As Long
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatestvolume As Double
    Dim tickergreatestincrease As String
    Dim tickergreatestdecrease As String
    Dim tickergreatestvolume As String
    
    ' Looping through each of the sheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Finding the last row of data
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Creating the table headers for requested values
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initializing variables
        summaryrow = 2
        greatestincrease = 0
        greatestdecrease = 0
        greatestvolume = 0
        tickergreatestincrease = ""
        tickergreatestdecrease = ""
        tickergreatestvolume = ""
        
        ' Looping through each row of data
        For i = 2 To lastrow
            ' Checking if the current ticker is the same as the previous ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Or i = 2 Then
                ' Storing the ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Storing the opening price
                openingrice = ws.Cells(i, 3).Value
                
                ' Resetting the total volume
                totalvolume = 0
            End If
            
            ' Storing the closing price
            closingrice = ws.Cells(i, 6).Value
            
            ' Calculating the yearly change
            yearlychange = closingrice - openingrice
            
            ' Calculating the percent change
            If openingrice <> 0 Then
                percentchange = yearlychange / openingrice * 100
            Else
                percentchange = 0
            End If
            
            ' Adding to total volume
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            
            ' Outputting results
            If ws.Cells(i + 1, 1).Value <> ticker Then
                
                ' Outputting results in summary table
                ws.Cells(summaryrow, 9).Value = ticker
                ws.Cells(summaryrow, 10).NumberFormat = "0.00"
                ws.Cells(summaryrow, 10).Value = yearlychange
                
                ws.Cells(summaryrow, 11).NumberFormat = "0.00\%"
                ws.Cells(summaryrow, 11).Value = percentchange
                
                ws.Cells(summaryrow, 12).NumberFormat = "0"
                ws.Cells(summaryrow, 12).Value = totalvolume
              
                ' Coloring cells based on yearly change. Red for decrease, green for increase
                If yearlychange > 0 Then
                    ws.Cells(summaryrow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlychange < 0 Then
                    ws.Cells(summaryrow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Checking for the greatest increase, decrease, and volume
                If percentchange > greatestincrease Then
                    greatestincrease = percentchange
                    tickergreatestincrease = ticker
                End If
                
                If percentchange < greatestdecrease Then
                    greatestdecrease = percentchange
                    tickergreatestdecrease = ticker
                End If
                
                If totalvolume > greatestvolume Then
                    greatestvolume = totalvolume
                    tickergreatestvolume = ticker
                End If
                
                ' Moving to the next row in the summary table
                summaryrow = summaryrow + 1
            End If
        Next i
        
          
        ' Outputting greatest increase, decrease, and volume
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = tickergreatestincrease
        ws.Cells(2, 17).Value = greatestincrease & "%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = tickergreatestdecrease
        ws.Cells(3, 17).Value = greatestdecrease & "%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = tickergreatestvolume
        ws.Cells(4, 17).Value = greatestvolume
        ws.Cells(4, 17).NumberFormat = "0.00E+00"

	' Auto Fitting columns
        ws.Columns("I:Q").AutoFit
        
    Next ws
    
End Sub
