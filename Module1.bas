Attribute VB_Name = "Module1"
Sub StockAnalysis()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTableRow As Long
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim tickerMaxPercentIncrease As String
    Dim tickerMaxPercentDecrease As String
    Dim tickerMaxTotalVolume As String
    ' Initialize variables for tracking the maximum values
    maxPercentIncrease = 0
    maxPercentDecrease = 0
    maxTotalVolume = 0
    tickerMaxPercentIncrease = ""
    tickerMaxPercentDecrease = ""
    tickerMaxTotalVolume = ""

    ' Loop through all worksheets in the workbook
    For Each ws In Worksheets
       
        ' Set up summary table headers
        ws.Cells(1, 14).Value = "Ticker"
        ws.Cells(1, 15).Value = "Yearly Change"
        ws.Cells(1, 16).Value = "Percent Change"
        ws.Cells(1, 17).Value = "Total Stock Volume"
       
        ' Initialize summary table row
        summaryTableRow = 2
       
        ' Initialize variables
        openingPrice = ws.Cells(2, 3).Value
        totalVolume = 0
       
        ' Find the last row of data in the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
       
        ' Loop through all rows in the current worksheet
        For i = 2 To lastRow
            ' Check if the current row's ticker symbol is different from the previous row
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Or i = 2 Then
                ' Set the ticker symbol
                ticker = ws.Cells(i, 1).Value
               
                closingPrice = ws.Cells(i - 1, 6).Value
                ' Calculate yearly change and percent change
                yearlyChange = closingPrice - openingPrice
                percentChange = yearlyChange / openingPrice
               
                ' Print the results in the summary table
                ws.Cells(summaryTableRow, 14).Value = ticker
                ws.Cells(summaryTableRow, 15).Value = yearlyChange
                ws.Cells(summaryTableRow, 16).Value = percentChange
                ws.Cells(summaryTableRow, 17).Value = totalVolume
               
                ' Format the percent change as a percentage
                ws.Cells(summaryTableRow, 11).NumberFormat = "0.00%"
               
                ' Conditional formatting for yearly change cell
                If yearlyChange >= 0 Then
                    ws.Cells(summaryTableRow, 15).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(summaryTableRow, 15).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                If percentChange >= 0 Then
                    ws.Cells(summaryTableRow, 16).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(summaryTableRow, 16).Interior.Color = RGB(255, 0, 0) ' Red
                End If
            ' Check if the current stock has the greatest percentage increase
                    If percentChange > maxPercentIncrease Then
                        maxPercentIncrease = percentChange
                        tickerMaxPercentIncrease = ticker
                    End If
                   
                                    ' Check if the current stock has the greatest percentage decrease
                    If percentChange < maxPercentDecrease Then
                        maxPercentDecrease = percentChange
                        tickerMaxPercentDecrease = ticker
                    End If
                   
                    ' Check if the current stock has the greatest total volume
                    If totalVolume > maxTotalVolume Then
                        maxTotalVolume = totalVolume
                        tickerMaxTotalVolume = ticker
                    End If
                ' Reset variables for the next ticker symbol
                openingPrice = ws.Cells(i, 3).Value
               
                totalVolume = 0
                ' Move to the next row in the summary table
                summaryTableRow = summaryTableRow + 1
            End If
           
            ' Accumulate the stock volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        Next i
       ' Output the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume
    ws.Cells(2, 20).Value = "Greatest % Increase"
    ws.Cells(3, 20).Value = "Greatest % Decrease"
    ws.Cells(4, 20).Value = "Greatest Total Volume"
    ws.Cells(1, 21).Value = "Ticker"
    ws.Cells(1, 22).Value = "Value"
   
    ws.Cells(2, 21).Value = tickerMaxPercentIncrease
    ws.Cells(3, 21).Value = tickerMaxPercentDecrease
    ws.Cells(4, 21).Value = tickerMaxTotalVolume
   
    ws.Cells(2, 22).Value = maxPercentIncrease
    ws.Cells(3, 22).Value = maxPercentDecrease
    ws.Cells(4, 22).Value = maxTotalVolume
   
    ' Format the percentage values as percentages
    ws.Cells(2, 21).NumberFormat = "0.00%"
    ws.Cells(3, 21).NumberFormat = "0.00%"

    Next ws
End Sub
