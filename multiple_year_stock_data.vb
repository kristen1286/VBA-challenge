Sub worksheetloop():

For Each ws In Worksheets                       'loop through all the worksheets
    Dim worksheetname As String
    worksheetname = ws.Name
                                                'set initial variables for calculations
    Dim ticker_name, max_tick_nam, min_tick_nam, maxvol_tick_nam As String
    ticker_name = ws.Cells(2, 1).Value
    
    Dim openprice, closeprice, yearchange, tickvol, yearchange_perc As Double
    Dim max_perc, min_perc, max_vol As Double
    max_perc = 0
    min_perc = 0
    max_vol = 0
    
    openprice = ws.Cells(2, 3).Value
    closeprice = 0
    yearchange = 0
    tickvol = 0
    yearchange_perc = 0
    
    Dim tablerow, lastrow As Long
    tablerow = 2                                        'set location of summary table
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row    'loop through all sheets to find last row that is not empty
    
    '####### create columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"


    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then     'iterates to see if the ticker is the same
            ticker_name = ws.Cells(i, 1).Value                          'ticker start cell
            closeprice = ws.Cells(i, 6).Value
            yearchange = closeprice - openprice
            
            'print in summary table
            ws.Range("I" & tablerow).Value = ticker_name
            ws.Range("J" & tablerow).Value = yearchange
            
                     
            
            'percent change part
            If openprice <> 0 Then
                yearchange_perc = (yearchange / openprice) * 100
            End If
            
            'ticker volume total calculation
            tickvol = tickvol + ws.Cells(i, 7).Value
            
    'print stuff
             'print percent year change in summary table
            ws.Range("K" & tablerow).Value = (CStr(yearchange_perc) & "%")
            'print ticker volume total
            ws.Range("L" & tablerow).Value = tickvol
            
    'color cells
            If (yearchange > 0 Or yearchange_perc > 0) Then
                ws.Range("J" & tablerow).Interior.ColorIndex = 4
                ws.Range("K" & tablerow).Interior.ColorIndex = 4
            ElseIf (yearchange <= 0 Or yearchange_perc > 0) Then
                ws.Range("J" & tablerow).Interior.ColorIndex = 3
                ws.Range("K" & tablerow).Interior.ColorIndex = 3
            End If
            
            
            'functionality calculations
            If (yearchange_perc > max_perc) Then
                max_perc = yearchange_perc
                max_tick_nam = ticker_name
            ElseIf (yearchange_perc < min_perc) Then
                min_perc = yearchange_perc
                min_tick_nam = ticker_name
            End If
            
            If (tickvol > max_vol) Then
                max_vol = tickvol
                maxvol_tick_nam = ticker_name
            End If
                
            '####################reset for next ticker
             'add counter
            tablerow = tablerow + 1
            
            'get new beginning price
            openprice = ws.Cells(i + 1, 3).Value
            
            'reset value
            yearchange_perc = 0
            tickvol = 0
        Else
            tickvol = tickvol + ws.Cells(i, 7).Value   'if in next ticker
            
        End If
       
        
    Next i
    
          '#############print functionality calculation
          ws.Range("Q2").Value = (CStr(max_perc) & "%")
          ws.Range("Q3").Value = (CStr(min_perc) & "%")
          ws.Range("P2").Value = max_tick_nam
          ws.Range("P3").Value = min_tick_nam
          ws.Range("P4").Value = maxvol_tick_nam
          ws.Range("Q4").Value = max_vol
          ws.Range("O2").Value = "Greatest % Increase"
          ws.Range("O3").Value = "Greatest % Decrease"
          ws.Range("O4").Value = "Greatest Total Volume"
Next ws
          
End Sub
