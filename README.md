# VBA-Challenge-
Create a script that loops through all the stocks for one year and outputs the following information:
The ticker symbol
Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
The total stock volume of the stock. The result should match the following image:
Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:
Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

---------
solution
---------
Sub Multiple_year_stock_data()

        '-----------------------------------------------------------
        'script that loops through all the stocks for one year
        '-----------------------------------------------------------
        
        'create a variable to hold ticker name
        Dim tickername As String
        
        'create a variable to hold the total volume
        Dim tickervolume As Double
        tickervolume = 0
        
        'create the counters
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        'Yearly Change: (Close Price at the end of a trading year - Open Price at the beginning of the trading year)
        
        'Percent change:((Close - Open)/Open)*100
        Dim openPrice As Double
        
        'Set openPrice
        open_price = Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        
        'create summary table headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest Percentage Increase"
        Cells(3, 15).Value = "Greatest Percentage Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        'Count the number of rows in the first column
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Looping through the rows
        
        For i = 2 To lastrow

            'looking for the point where the next value in the sheet is different than that of the current cell
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              'extracting the ticker name
              tickername = Cells(i, 1).Value

              'calculating the volume for every ticker
              tickervolume = tickervolume + Cells(i, 7).Value

              'Print the ticker name in the summary table
              Range("I" & summary_ticker_row).Value = tickername

              'Print the trade volume for each ticker in the summary table
              Range("L" & summary_ticker_row).Value = tickervolume

              'adding closing price at the end of the year
              closePrice = Cells(i, 6).Value

              'Calculate yearly change from end of the year to start of the year
               yearlyChange = (closePrice - openPrice)
              
              'Print the yearly change for each ticker in the summary table
              Range("J" & summary_ticker_row).Value = yearlyChange

              'set value to 0 to avoid divisibility by zero
                If openPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openPrice
                End If

              'Print the yearly change for each ticker in the summary table
              Range("K" & summary_ticker_row).Value = percentChange
              Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter and add one to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of trade to zero
              tickervolume = 0

              'Reset the opening price
              openPrice = Cells(i + 1, 3)
            
            Else
              
               'Add the volume
              tickervolume = tickervolume + Cells(i, 7).Value

            
            End If
        
        Next i

    'Conditional formatting that will highlight positive change in green and negative change in red
    
    'find the last row of the table
    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code for yearly change
        For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

    'Determine the max and min values
    
        For i = 2 To lastrow_summary_table
            'maximum percent change
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            'minimum percent change
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
            'maximum volume of trade
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i
        
End Sub
--------------------------------------------------------
'I started with the code for one sheet
'Now expand it for all the worksheets

Sub Multiple_year_stock_data()

        '-----------------------------------------------------------
        'script that loops through all the sheets
        '-----------------------------------------------------------
      For each ws In Worksheets
        
        'create a variable to hold ticker name
        Dim tickername As String
        
        'create a variable to hold the total volume
        Dim tickervolume As Double
        tickervolume = 0
        
        'create the counters
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        'Yearly Change: (Close Price at the end of a trading year - Open Price at the beginning of the trading year)
        
        'Percent change:((Close - Open)/Open)*100
        Dim openPrice As Double
        
        'Set openPrice
        open_price = ws.Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        
        'create summary table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest Percentage Increase"
        ws.Cells(3, 15).Value = "Greatest Percentage Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Count the number of rows in the first column
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Looping through the rows
        
        For i = 2 To lastrow

            'looking for the point where the next value in the sheet is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'extracting the ticker name
              tickername = ws.Cells(i, 1).Value

              'calculating the volume for every ticker
              tickervolume = tickervolume + ws.Cells(i, 7).Value

              'Print the ticker name in the summary table
              ws.Range("I" & summary_ticker_row).Value = tickername

              'Print the trade volume for each ticker in the summary table
              ws.Range("L" & summary_ticker_row).Value = tickervolume

              'adding closing price at the end of the year
              closePrice = ws.Cells(i, 6).Value

              'Calculate yearly change from end of the year to start of the year
               yearlyChange = (closePrice - openPrice)
              
              'Print the yearly change for each ticker in the summary table
              ws.Range("J" & summary_ticker_row).Value = yearlyChange

              'set value to 0 to avoid divisibility by zero
                If openPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openPrice
                End If

              'Print the yearly change for each ticker in the summary table
              ws.Range("K" & summary_ticker_row).Value = percentChange
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter and add one to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of trade to zero
              tickervolume = 0

              'Reset the opening price
              openPrice = ws.Cells(i + 1, 3)
            
            Else
              
               'Add the volume
              tickervolume = tickervolume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    'Conditional formatting that will highlight positive change in green and negative change in red
    
    'find the last row of the table
    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code for yearly change
        For i = 2 To lastrow_summary_table
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

    'Determine the max and min values
    
        For i = 2 To lastrow_summary_table
            'maximum percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'minimum percent change
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'maximum volume of trade
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
        
End Sub



