# VBA-challenge
Sub MultipleYearStockData()
    Dim i As Long, rowCount As Long 'used as counters for looping through rows
    Dim ws As Worksheet 'stores the worksheet where the data is located
    Dim ticker As String 'stores the stock ticker of the current row
    Dim outRow As Long, start_row As Long 'tracks rows in the summary table
    Dim opening As Double, closing As Double 'store the stock's opening and closing prices
    Dim year_ As Integer, quarter As Integer 'store the year and quarter information
    Dim open_date As Date, close_date As Date 'track dates of opening and closing prices
    Dim quarterly_change As Double, percent_change As Double 'store calculated changes in stock prices
    Dim total_stock_volume As LongLong 'stores the total volume of stock traded within a quarter
    Dim summary_table_row As Long 'tracks the current row in the summary table where results are outputted
    
    ' Set the worksheet containing the data- which stores
    For Each ws In ThisWorkbook.Worksheets
        If ws.Index > 4 Then Exit For 'this command will end the loop at
    
    ' Determine the last row with data
    rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Initialize the summary table headers, which will set up column heaers ofr the summary output in Columns  I through L
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Set the starting row for summary output- initialize to start outputting summary data from row 2
    summary_table_row = 2
    
    '"For Loop" through all stock data
    For i = 2 To rowCount
        ' Get the current ticker, open date, and opening price
        ticker = ws.Cells(i, 1).Value 'stores the current ticker symbol
        open_date = ws.Cells(i, 2).Value 'stores the date of the stock record in this row
        opening = ws.Cells(i, 3).Value 'stores the opening price of the stock in this row
        
        'calculating quarter and year from open_date
        quarter = Application.WorksheetFunction.RoundUp(Month(open_date) / 3, 0) 'calcuates the quarter by dividing the month by 3 and rounding up (so month 1-3 = Q1, etc.)
        year_ = Year(open_date) 'extracts the year
        
        ' Reset the total stock volume for each new ticker and quarter
        total_stock_volume = 0
        
    ' "While loop" to accumulate data for the current quarter
    Do While i <= rowCount And ws.Cells(i, 1).Value = ticker _
        And Year(ws.Cells(i, 2).Value) = year_ _
        And Application.WorksheetFunction.RoundUp(Month(ws.Cells(i, 2).Value) / 3, 0) = quarter
            
    ' Accumulate total stock volume and update the closing price
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        closing = ws.Cells(i, 6).Value
        i = i + 1
        Loop
        
        ' Move back one row to correctly end the loop
        i = i - 1
        
    ' Calculate the quarterly change and percent change
        quarterly_change = closing - opening 'calculates the change in price from the quarter's start to end
        If opening <> 0 Then 'checking that opening price is not 0 to avoid errors in division
            percent_change = (quarterly_change / opening) * 100 'calculates percent change in price
        Else
            percent_change = 0
        End If
        
        ' Output the results to the summary table
        ws.Cells(summary_table_row, 9).Value = ticker
        ws.Cells(summary_table_row, 10).Value = Round(quarterly_change, 2)
        ws.Cells(summary_table_row, 11).Value = Round(percent_change, 2) & "%"
        ws.Cells(summary_table_row, 12).Value = total_stock_volume
        
        'Conditional Formatting
        If ws.Cells(summary_table_row, 10).Value > 0 Then
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4 ' Green for positive change
            ElseIf ws.Cells(summary_table_row, 10).Value < 0 Then
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3 ' Red for negative change
        End If
        ' Move to the next row in the summary table
        summary_table_row = summary_table_row + 1
        
    Next i

    ' Format the percent change column to a percent
    ws.Columns("K").NumberFormat = "0.00%"
    Next ws
    
End Sub
