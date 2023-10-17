VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Loop()

    For Each ws In Worksheets
    
        Dim tickername As String
        Dim tickervolume As Double
        tickervolume = 0
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        'Yearly change is: (Close Price - Open Price)
        'Percent change is a simple percent change -->((Close - Open)/Open)*100
        
        Dim open_price As Double
        
        'Set initial open_price. Other opening prices will be determined in the conditional loop.
        
        open_price = ws.Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        'Create header titles
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Count the number of rows in the first column.
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through the rows by the ticker names

        For i = 2 To LastRow

            'Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'Set the ticker name
              tickername = ws.Cells(i, 1).Value

              'Add the volume of trade
              tickervolume = tickervolume + ws.Cells(i, 7).Value

              'Print the ticker name in the summary table
              ws.Range("I" & summary_ticker_row).Value = tickername

              'Print the trade volume for each ticker in the summary table
              ws.Range("L" & summary_ticker_row).Value = tickervolume

              'Now collect information about closing price
              close_price = ws.Cells(i, 6).Value

              'Calculate yearly change
               yearly_change = (close_price - open_price)
              
              'Print the yearly change for each ticker in the summary table
              ws.Range("J" & summary_ticker_row).Value = yearly_change

              'Check for the non-divisibilty condition when calculating the percent change
                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If

              'Print the yearly change for each ticker in the summary table
              ws.Range("K" & summary_ticker_row).Value = percent_change
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter. Add one to the summary_ticker_row
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume of trade to zero
              tickervolume = 0

              'Reset the opening price
              open_price = ws.Cells(i + 1, 3)
            
            Else
              
               'Add the volume of trade
              tickervolume = tickervolume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Conditional format by positive or negative values (red = neg and green = pos)
        
        For i = 2 To lastrow_summary_table
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

    'Find max and min values in column "Percent Change" and max for column "Total Stock Volume"
    'Then grab the ticker name, and the corresponding values for the percent change and total volume of trade for that ticker
    '
        For i = 2 To lastrow_summary_table
            'Find the maximum percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = Cells(i, 9).Value
                ws.Cells(2, 17).Value = Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Find the minimum percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = Cells(i, 9).Value
                ws.Cells(3, 17).Value = Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        
End Sub
