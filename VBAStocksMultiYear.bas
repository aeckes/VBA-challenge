Sub VBAStocks()

Dim first_price, last_price, volume_agg, percent_change, yearly_change, lrg_pct_incr, lrg_pct_decr, lrg_total_vol As Double
Dim ticker As String
Dim i, j, ticker_counter As Integer
Dim LastUsedRow As Long
Dim ws As Worksheet

'do same for each ws
For Each ws In ActiveWorkbook.Worksheets
    
    'delete before submission
    ws.Range("J:L").ClearFormats

    'last row number on ws
    LastUsedRow = ws.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
    
    'define column headers / row labels
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    

    
    'get initial ticker symbol
    ticker = ws.Cells(2, 1).Value
    ticker_counter = 2
    
    'reset values
    volume_agg = 0
    percent_change = 0
    yearly_change = 0
    lrg_pct_incr = 0
    lrg_pct_decr = 0
    lrg_total_vol = 0
    
    'print ticker symbol
    ws.Cells(ticker_counter, 9).Value = ticker
    
    'get initial ticker open price
    first_price = ws.Cells(2, 3).Value
    
    'cycle thru rows looking for next symbol
    For i = 2 To LastUsedRow + 1

        If ws.Cells(i, 1).Value <> ticker Then
        
        last_price = ws.Cells(i - 1, 6).Value
        yearly_change = last_price - first_price
        
        If last_price = 0 Or first_price = 0 Then
            percent_change = 0
        Else
            percent_change = last_price / first_price - 1
        End If
                
        ws.Cells(ticker_counter, 10).Value = yearly_change
        ws.Cells(ticker_counter, 11).Value = percent_change
        ws.Cells(ticker_counter, 12).Value = volume_agg
        
        If yearly_change < 0 Then
            ws.Cells(ticker_counter, 10).Interior.Color = 255
        Else
            ws.Cells(ticker_counter, 10).Interior.Color = 5296274
        End If
        
        If percent_change > lrg_pct_incr Then
            lrg_pct_incr = percent_change
            ws.Cells(2, 17).Value = lrg_pct_incr
            ws.Cells(2, 16).Value = ticker
        Else
        If percent_change < lrg_pct_decr Then
            lrg_pct_decr = percent_change
            ws.Cells(3, 17).Value = lrg_pct_decr
            ws.Cells(3, 16).Value = ticker
        Else
        If volume_agg > lrg_total_vol Then
            lrg_total_vol = volume_agg
            ws.Cells(4, 17).Value = lrg_total_vol
            ws.Cells(4, 16).Value = ticker
        End If
        End If
        End If
        
        'Get next ticker first price
        first_price = ws.Cells(i, 3).Value
        ticker = ws.Cells(i, 1).Value
        ticker_counter = ticker_counter + 1
        ws.Cells(ticker_counter, 9).Value = ticker
        
        volume_agg = 0
        percent_change = 0
        yearly_change = 0
        
        End If

    'aggregate volume each row per ticker
    volume_agg = volume_agg + ws.Cells(i, 7).Value

    Next i
   
    'apply formatting
    ws.Range("J:J").NumberFormat = "#,##0.00"
    ws.Range("K:K").NumberFormat = "#.00%"
    ws.Range("L:L").NumberFormat = "#,##0"
    
    ws.Cells(2, 17).NumberFormat = "#.00%"
    ws.Cells(3, 17).NumberFormat = "#.00%"
    ws.Cells(4, 17).NumberFormat = "#,##0"
    
Next ws

End Sub


