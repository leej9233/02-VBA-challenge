Sub stock():
    'define variables
    Dim ticker As String
    Dim yearly_change As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim percent_change As Double
    Dim total_volume As LongLong
    Dim ws_count As Integer
    Dim year As Integer
    
    'Set ws_count to count the number of worksheets
    ws_count = ActiveWorkbook.Worksheets.Count
    
    For year = 1 To ws_count
        
        'set variables
        endr = Cells(Rows.Count, 1).End(xlUp).Row
        output = 2
        total_volume = 0
        
        'Activate worksheet
        Worksheets(year).Activate
        
        
        'start loop
        For i = 2 To endr
        
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                'header for analysis
                Cells(1, 9).Value = "Ticker"
                Cells(1, 10).Value = "Yearly Change"
                Cells(1, 11).Value = "Percent Change"
                Cells(1, 12).Value = "Total Stock Volume"
                
                'ticker name in report
                Cells(output, 9).Value = Cells(i, 1).Value
                
                'close price value
                close_price = Cells(i, 6).Value
                
                'yearly change
                yearly_change = close_price - open_price
                
                'change yearly change format
                Cells(output, 10).NumberFormat = "0.00"
                
                
                'print yearly change in report
                Cells(output, 10).Value = yearly_change
                
                'color for yearly change
                If yearly_change > 0 Then
                    Cells(output, 10).Interior.Color = vbGreen
                Else
                    Cells(output, 10).Interior.Color = vbRed
                End If
                
                'percent change
                If open_price <> 0 Then
                    percent_change = (yearly_change / open_price)
                Else
                    percent_change = yearly_change
                End If
                
                'change percent change format
                Cells(output, 11).NumberFormat = "0.00%"
                
                'print percent change in report
                Cells(output, 11).Value = percent_change
                
                'add to total value anr print in report
                total_volume = total_volume + Cells(i, 7).Value
                Cells(output, 12).Value = total_volume
                
                'reset
                total_volume = 0
                
                'go to the next row
                output = output + 1
            ElseIf Cells(i, 2).Value = "20180102" Or Cells(i, 2).Value = "20190102" Or Cells(i, 2).Value = "20200102" Then
            
                'open price value
                open_price = Cells(i, 3).Value
                
                'keep adding total volume
                total_volume = total_volume + Cells(i, 7).Value
                
            Else
                'keep adding total volume
                total_volume = total_volume + Cells(i, 7).Value
                
            End If
            
        Next i
        
        'header
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        'find min/max percentage
        g_percent_inc = Application.WorksheetFunction.Max(Range("K:K"))
        g_percent_dec = Application.WorksheetFunction.Min(Range("K:K"))
        
        'print in report
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(2, 17).Value = g_percent_inc
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(3, 17).Value = g_percent_dec
        
        'change format into percentage
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).NumberFormat = "0.00%"
        For i = 2 To endr
            If Cells(i, 11).Value = g_percent_inc Then
                max_name = Cells(i, 9).Value
            ElseIf Cells(i, 11).Value = g_percent_dec Then
                min_name = Cells(i, 9).Value
            End If
        Next i
        
        'print ticker name in report
        Cells(2, 16).Value = max_name
        Cells(3, 16).Value = min_name
        
        'find max volume
        g_total_volume = Application.WorksheetFunction.Max(Range("L:L"))
        
        'print in report
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(4, 17).Value = g_total_volume
        For i = 2 To endr
            If Cells(i, 12).Value = g_total_volume Then
                g_total_volume_name = Cells(i, 9).Value
            End If
        Next i
        
        'print ticker name in report
        Cells(4, 16).Value = g_total_volume_name
    
    Next year
    
End Sub
