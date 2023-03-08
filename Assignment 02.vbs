Sub multistockdata()

'State all variables used

Dim lastrow As Long
Dim countrows As Long
Dim i As Long
Dim yearlychange As Double
Dim stockvolume As Double
Dim max_change As Double
Dim min_change As Double
Dim max_vol As Double

 For Each ws In Worksheets
 
    'Determine titles for the summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest Percentage Increase Change"
    ws.Cells(3, 15).Value = "Greatest Percentage Decrease Change"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
     
    'Determine total number of rows of the data
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'location of ticker names for the summary table
    countrows = 2
    
    'Set initial variable for holding the yearly change for the ticker
    yearlychange = 0
    
    'Set initial variable for holding the stock volume
    stockvolume = 0
    
    'Set initial value of the open price for the first ticker
     openprice = ws.Cells(2, 3).Value
    
    'Start looping for all tickers

        For i = 2 To lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Set the data for ticker name and all related values
            ticker_name = ws.Cells(i, 1).Value
            closeprice = ws.Cells(i, 6).Value
            
           'Calculate Yearlychange
            yearlychange = closeprice - openprice
           
                
           'Calculate Percentage change and change format to %
           'As value divided by zero will give N/A, hence correction is needed
            
                If yearlychange = 0 Or openprice = 0 Then
                  
                ws.Cells(countrows, 11).Value = 0
                  
                Else
                  
                ws.Cells(countrows, 11).Value = Format(yearlychange / openprice, "0.00%")
                  
                End If
                
                percent_change = ws.Cells(countrows, 11)
            
           'Maximum Percentage Change
            
                If percent_change < max_change Then
                max_change = max_change
                                       
                ElseIf percent_change > max_change Then
                max_change = percent_change
                                       
                ws.Cells(2, 17).Value = max_change
                max_ticker_name = ticker_name
                                       
                End If
                
                 
            'Minimum Percentage Change
                
                If percent_change < min_change Then
                min_change = percent_change
                                          
                ElseIf percent_change > min_change Then
                min_change = min_change
                                          
                ws.Cells(3, 17).Value = min_change
                min_ticker_name = ticker_name
                                          
                End If
                 
            
            'Calculate total ticker volume
            stockvolume = stockvolume + ws.Cells(i, 7).Value
            
            'Greates Total Stock Volume
                
                If stockvolume < max_vol Then
                max_vol = max_vol
                                          
                ElseIf stockvolume > max_vol Then
                max_vol = stockvolume
                                          
                ws.Cells(4, 17).Value = max_vol
                max_volume_name = ticker_name
                                          
                End If
            
            'Location of summary tables for ticker name and all related values
            ws.Range("I" & countrows).Value = ticker_name
            ws.Range("J" & countrows).Value = yearlychange
            ws.Range("L" & countrows).Value = stockvolume
            ws.Range("P2").Value = max_ticker_name
            ws.Range("P3").Value = min_ticker_name
             ws.Range("P4").Value = max_volume_name
           
            'Conditional formatting for positive or negative change
                If yearlychange < 0 Then
                ws.Cells(countrows, 10).Interior.ColorIndex = 3
               
                ElseIf yearlychange >= 0 Then
                ws.Cells(countrows, 10).Interior.ColorIndex = 4
               
                End If
                
                If yearlychange < 0 Then
                ws.Cells(countrows, 11).Interior.ColorIndex = 3
                    
                ElseIf yearlychange >= 0 Then
                ws.Cells(countrows, 11).Interior.ColorIndex = 4
                
                End If
                
            'Add new row when work with new tickers
            countrows = countrows + 1
            
            'Reset back the values as working with new tickers
            closeprice = 0
            stockvolume = 0
            
            'Set new ticker open price for new ticker
            openprice = ws.Cells(i + 1, 3).Value
            
            Else
            
           'Combined all price differences for same type of ticker
            yearlychange = yearlychange + (closeprice - openprice)
            stockvolume = stockvolume + ws.Cells(i, 7).Value
            
            End If
        
    Next i
    
Next ws

End Sub




