Sub CalculateStock()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryRow As Long
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row of the stock data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Set the headers for the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize summary row and variables for tracking maximum values
        summaryRow = 2
        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxTotalVolume = 0
        maxPercentIncreaseTicker = ""
        maxPercentDecreaseTicker = ""
        maxTotalVolumeTicker = ""
        
        ' Loop through the stock data
        For i = 2 To lastRow ' Assuming the data starts from row 2
            ' Check if the ticker symbol has changed
            If ws.Cells(i, 1).Value <> ticker Then
                ' Save the new ticker symbol and opening price
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
                totalVolume = 0 ' Reset total volume for the new ticker
            End If
            
            ' Accumulate the total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if it's the last record for the current ticker
            If ws.Cells(i + 1, 1).Value <> ticker Then
                ' Save the closing price and calculate yearly change and percent change
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice
                
                ' Output the summary information
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Format the percent change column as percentage
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                
                ' Format positive change in green and negative change in red
                If yearlyChange >= 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
                Else
                    ws.Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If
                
                ' Check if the current stock has the greatest percent increase
                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    maxPercentIncreaseTicker = ticker
                End If
                
                ' Check if the current stock has the greatest percent decrease
                If percentChange < maxPercentDecrease Then
                    maxPercentDecrease = percentChange
                    maxPercentDecreaseTicker = ticker
                End If
                
                ' Check if the current stock has the greatest total volume
                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    maxTotalVolumeTicker = ticker
                End If
                
                ' Move to the next row in the summary table
                summaryRow = summaryRow + 1
            End If
        Next i
        
        ' Output the stock with the greatest percent increase, greatest percent decrease, and greatest total volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 16).Value = maxPercentIncreaseTicker
        ws.Cells(2, 17).Value = maxPercentIncrease
        ws.Cells(3, 16).Value = maxPercentDecreaseTicker
        ws.Cells(3, 17).Value = maxPercentDecrease
        ws.Cells(4, 16).Value = maxTotalVolumeTicker
        ws.Cells(4, 17).Value = maxTotalVolume
        
        ' Format the percentage columns in the summary table
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ' Autofit the columns in the summary table
        ws.Columns("I:L").AutoFit
        ws.Columns("O:R").AutoFit
    Next ws
End Sub

