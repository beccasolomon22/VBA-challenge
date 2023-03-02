Attribute VB_Name = "Module1"
Sub stock_vol_all()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
        'find last row of sheet
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'label columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        'set all variables
        ticker = ws.Cells(2, 1).Value
        ticker_row = 2
        open_value = ws.Cells(2, 3).Value
        close_value = 0
        total_stock = ws.Cells(2, 7).Value
    
        'insert first ticker
        ws.Cells(2, 9).Value = ticker
        
        'for loop that iterates through all tickers
    
        For i = 2 To last_row
            'if the ticker we are on is different than the next we move to next ticker
            If (ticker <> ws.Cells(i + 1, 1).Value) <> False Then
                ticker = ws.Cells(i + 1, 1).Value
                ws.Cells(ticker_row + 1, 9).Value = ticker
                close_value = ws.Cells(i, 6).Value
                
                'determine yearly change
                ws.Cells(ticker_row, 10).Value = close_value - open_value
                ws.Cells(ticker_row, 10).NumberFormat = "0.00"
                
                'format yearly change. red if negative, green if positive and purple if zero
                If ws.Cells(ticker_row, 10).Value > 0 Then
                    ws.Cells(ticker_row, 10).Interior.ColorIndex = 4
                
                ElseIf ws.Cells(ticker_row, 10).Value < 0 Then
                    ws.Cells(ticker_row, 10).Interior.ColorIndex = 3
                
                Else
                    ws.Cells(ticker_row, 10).Interior.ColorIndex = 17
                End If
            
                'percent change
                ws.Cells(ticker_row, 11).Value = (ws.Cells(ticker_row, 10).Value / open_value)
                ws.Cells(ticker_row, 11).Value = FormatPercent(ws.Cells(ticker_row, 11))
                
                'set up for next loop
                ticker_row = ticker_row + 1
                open_value = ws.Cells(i + 1, 3).Value
                total_stock = ws.Cells(i + 1, 7).Value
            Else
                'if tickers are the same add the total stock value to the total stock volume
                ws.Cells(ticker_row, 12).Value = total_stock
                total_stock = total_stock + ws.Cells(i + 1, 7).Value
            
            End If
        Next i
            
        'add the rows for the greatest ""
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        'set variable for greatest ""
        highest_percent = ws.Cells(2, 11).Value
        lowest_percent = ws.Cells(2, 11).Value
        greatest_vol = ws.Cells(2, 12).Value
    
        highest_ticker = ws.Cells(2, 9).Value
        lowest_ticker = ws.Cells(2, 9).Value
        total_ticker = ws.Cells(2, 12).Value
    
        'set columns for greatest ""
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'find end of the new ticker column
        end_column9 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        'for loop to determine greatest percent increase, greatest percent decrease, and greatest total volume
        For i = 3 To end_column9
            'check and update greatest percent increase or greatest percent decrease
            If highest_percent < ws.Cells(i, 11).Value Then
                highest_percent = ws.Cells(i, 11).Value
                highest_ticker = ws.Cells(i, 9).Value
            
            ElseIf lowest_percent > ws.Cells(i, 11).Value Then
                lowest_percent = ws.Cells(i, 11).Value
                lowest_ticker = ws.Cells(i, 9).Value
            Else
            
            End If
        
            'check and update greatest total volume
            If greatest_vol < ws.Cells(i, 12).Value Then
                greatest_vol = ws.Cells(i, 12).Value
                total_ticker = ws.Cells(i, 9).Value
            Else
                'i didn't know what to add for an else here
            End If
        
        Next i
            
        'functionality for greatest increase/decrease/total
        ws.Cells(2, 16).Value = highest_ticker
        ws.Cells(3, 16).Value = lowest_ticker
        ws.Cells(4, 16).Value = total_ticker
        ws.Cells(2, 17).Value = highest_percent
        ws.Cells(3, 17).Value = lowest_percent
        ws.Cells(4, 17).Value = greatest_vol
        
        'formats those to percent
        ws.Cells(2, 17).Value = FormatPercent(ws.Cells(2, 17))
        ws.Cells(3, 17).Value = FormatPercent(ws.Cells(3, 17))
        
    Next ws

        
End Sub


Sub stock_vol_current()

    'in case you are crazy and want to only run it on one page and not all... here you go
    
    'find last row of sheet
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'label columns
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'set all variables for first loop
    ticker = Cells(2, 1).Value
    ticker_row = 2
    open_value = Cells(2, 3).Value
    close_value = 0
    total_stock = Cells(2, 7).Value
    
    'insert first ticker
    Cells(2, 9).Value = ticker
    
    'for loop that iterates through all tickers
    
    For i = 2 To last_row
        'if the ticker we are on is different than the next we move to next ticker
        
        If (ticker <> Cells(i + 1, 1).Value) <> False Then
            ticker = Cells(i + 1, 1).Value
            Cells(ticker_row + 1, 9).Value = ticker
            close_value = Cells(i, 6).Value
            
            'determine yearly change
            Cells(ticker_row, 10).Value = close_value - open_value
            ActiveCell.Select
            Selection.Value = Format(ActiveCell, "#.00")
            
            'format yearly change. red if negative, green if positive and purple if zero
            If Cells(ticker_row, 10).Value > 0 Then
                Cells(ticker_row, 10).Interior.ColorIndex = 4
                
            ElseIf Cells(ticker_row, 10).Value < 0 Then
                Cells(ticker_row, 10).Interior.ColorIndex = 3
                
            Else
                Cells(ticker_row, 10).Interior.ColorIndex = 17
            End If
            
            'percent change
            Cells(ticker_row, 11).Value = (Cells(ticker_row, 10).Value / open_value)
            Cells(ticker_row, 10).NumberFormat = "0.00"
            
            'set up for next loop
            ticker_row = ticker_row + 1
            open_value = Cells(i + 1, 3).Value
            total_stock = Cells(i + 1, 7).Value
            
        Else
            'if tickers are the same add the total stock value to the total stock volume
            Cells(ticker_row, 12).Value = total_stock
            total_stock = total_stock + Cells(i + 1, 7).Value
            
        End If
        
    Next i
            
    'add the rows for the greatest ""
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'set variable for greatest ""
    highest_percent = Cells(2, 11).Value
    lowest_percent = Cells(2, 11).Value
    greatest_vol = Cells(2, 12).Value
    
    highest_ticker = Cells(2, 9).Value
    lowest_ticker = Cells(2, 9).Value
    total_ticker = Cells(2, 12).Value
    
    'set columns for greatest ""
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'find end of the new ticker column
    end_column9 = Cells(Rows.Count, 9).End(xlUp).Row
    
    'for loop to determine greatest percent increase, greatest percent decrease, and greatest total volume
    For i = 3 To end_column9
        'check and update greatest percent increase or greatest percent decrease
        If highest_percent < Cells(i, 11).Value Then
            highest_percent = Cells(i, 11).Value
            highest_ticker = Cells(i, 9).Value
            
        ElseIf lowest_percent > Cells(i, 11).Value Then
                lowest_percent = Cells(i, 11).Value
                lowest_ticker = Cells(i, 9).Value
            Else
            
        End If
        
        'check and update greatest total volume
        If greatest_vol < Cells(i, 12).Value Then
            greatest_vol = Cells(i, 12).Value
            total_ticker = Cells(i, 9).Value
        Else
            'i didn't know what to add for an else here
        End If
        
    Next i
            
    'functionality for greatest increase/decrease/total
    Cells(2, 16).Value = highest_ticker
    Cells(3, 16).Value = lowest_ticker
    Cells(4, 16).Value = total_ticker
    Cells(2, 17).Value = highest_percent
    Cells(3, 17).Value = lowest_percent
    Cells(4, 17).Value = greatest_vol
        
    'formats those to percent
    Cells(2, 17).Value = FormatPercent(Cells(2, 17))
    Cells(3, 17).Value = FormatPercent(Cells(3, 17))


End Sub

