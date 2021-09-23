Attribute VB_Name = "Module1"
Sub multiYear():
    For Each ws In Worksheets
    
        'Variable to hold ticker
        Dim ticker, tickerGreatInc, tickerGreatDec, tickerGreatTotal As String
        'variable to hold last row
        Dim lastRow As Long
        'variable to hold total stock volume
        Dim totalVolume, greatest_totalVol As LongLong
        totalVolume = 0
        'variable for open and close, yearly and percent changes
        Dim tickerOpen, tickerClose, yearlyChange, percentChange, greatest_per_inc, greatest_per_dec As Double
        
        'add the word Ticker in cell "I1" on sheet1
        ws.Range("I1").Value = "Ticker"
                
        'add the word Yearly Change in cell "J1"
        ws.Range("J1").Value = "Yearly Change"
        
        'add the word Percent Change in cell "K1"
        ws.Range("K1").Value = "Percent Change"
        
        'add the words Total Stock Volume in cell "L1"
        ws.Range("L1").Value = "Total Stock Volume"
                                        
        'summary table row
        Dim summaryTableRow As Integer
        summaryTableRow = 2 'starts as row 2 in summary table
        
        'count number of rows
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'loop through all ticker rows
        For i = 2 To lastRow
            
            If (ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value) Then
            
                'set (reset) the tickerOpen
                tickerOpen = ws.Range("C" & i).Value
            
            End If
            
            'check to see if we are still within the same ticker
            'if not, do the following
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
        
                'set (reset) the ticker and tickerClose
                ticker = ws.Range("A" & i).Value
                tickerClose = ws.Range("F" & i).Value
                
                'subtract tickerOpen from tickerClose
                yearlyChange = tickerClose - tickerOpen
                
                'find percentChange value, if tickerOpen is zero, then percentChange = 0
                If tickerOpen > 0 Then
                    percentChange = yearlyChange / tickerOpen
                
                Else
                    percentChange = 0
                    
                End If
                
                'add yearlyChange to column J on current summary table row
                ws.Range("J" & summaryTableRow).Value = yearlyChange
                
                'add percentChange value to column K on current summary table row
                ws.Range("K" & summaryTableRow).Value = percentChange
                
                'add to the total stock volume one last time before the change in ticker
                totalVolume = totalVolume + ws.Range("G" & i).Value
            
                'add the values to the summary table
                'add the ticker to column I on the current summary table row
                ws.Range("I" & summaryTableRow).Value = ticker
                'add the total stock volume to column L on the current summary table row
                ws.Range("L" & summaryTableRow).Value = totalVolume
                                
                'color yearlyChange cells based on positive and negative changes
                If yearlyChange > 0 Then
                
                    'positive change in green
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
                
                ElseIf yearlyChange < 0 Then
                    'negative changes in red
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
                    
                ElseIf yearlyChange = 0 Then
                    'keep cell white
                    ws.Cells(summaryTableRow, 10).Interior.ColorIndex = 2
                
                End If
                                
                'once the summary table is populated, then add one to the summary row count
                summaryTableRow = summaryTableRow + 1
                'then reset the total stock volume
                totalVolume = 0
                                                
            Else
                'if we're in the same ticker, add on to the running total
                totalVolume = totalVolume + ws.Range("G" & i).Value
                     
            End If
                                       
        Next i
              
        'apply percent style to column K on current summary table row
        ws.Columns("K").NumberFormat = "0.00%"
        
        'add greatest percent increase and decrease and total volume text in cells "O2" "O3" and "O4"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'add the work Ticker" to "P1" on sheet1
        ws.Range("P1").Value = "Ticker"
        
        'add the word Value in cell "Q1" on sheet1
        ws.Range("Q1").Value = "Value"
        
        'find greatest percent increase
        greatest_per_inc = WorksheetFunction.Max(ws.Range("K2:K" & (summaryTableRow - 1)))
        tickerGreatInc = WorksheetFunction.Match(greatest_per_inc, ws.Range("K2:K" & (summaryTableRow - 1)), 0)
        ws.Range("P2").Value = ws.Range("I" & tickerGreatInc + 1).Value
        ws.Range("Q2").Value = greatest_per_inc
        
        'find greatest percent decrease
        greatest_per_dec = WorksheetFunction.Min(ws.Range("K2:K" & (summaryTableRow - 1)))
        tickerGreatDec = WorksheetFunction.Match(greatest_per_dec, ws.Range("K2:K" & (summaryTableRow - 1)), 0)
        ws.Range("P3").Value = ws.Range("I" & tickerGreatDec + 1).Value
        ws.Range("Q3").Value = greatest_per_dec
    
        'find greatest total volume
        greatest_totalVol = WorksheetFunction.Max(ws.Range("L2:L" & (summaryTableRow - 1)))
        tickerGreatTotal = WorksheetFunction.Match(greatest_totalVol, ws.Range("L2:L" & (summaryTableRow - 1)), 0)
        ws.Range("P4").Value = ws.Range("I" & tickerGreatTotal + 1).Value
        ws.Range("Q4").Value = greatest_totalVol
        
        'apply percent styles
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        'autofit columns
        ws.Range("A:Q").Columns.AutoFit
                                       
    Next ws
 
End Sub
