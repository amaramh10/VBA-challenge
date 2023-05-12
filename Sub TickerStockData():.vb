Sub TickerStockData():

    'Loop through worksheets
    For Each ws In Worksheets
    
        'Set variable for ticker symbols
        Dim TickerSymbol As String
        
        'Set Variables for Open & Close
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        
        'Set Variable for the yearly change
        Dim YearlyChange As Double
        YearlyChange = 0
        
        'Set Variable for percentage change
        Dim PercentChange As Double
        PercentChange = 0
        
        'Set Variable for total stock volume
        Dim TotalStockVol As Double
        TotalStockVol = 0
        
        'Keep Track of summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Set Ticker Start Row Variable
        Dim StartRow As Long
        StartRow = 2
        
        
        'Identify Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Add Column Headers for Summary Table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Loop through different ticker symbols
        For i = 2 To LastRow
        
            'Check if we are in the same ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Set the Ticker
                Ticker = ws.Cells(i, 1).Value
                
                'Print the Ticker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                'Set the Open & Close Price & Yearly Change
                OpenPrice = ws.Range("C" & StartRow).Value
                ClosePrice = ws.Range("F" & i).Value
                YearlyChange = (ClosePrice - OpenPrice)
                
                'Print the Yearly Change in the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                
                'Set the Percent Change
                 PercentChange = YearlyChange / ClosePrice
               
                'Print the Percent Change in the Summary Table & Change to percentage
                ws.Range("K" & Summary_Table_Row).Value = PercentChange
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                'Set & add to the Total Stock Volume
                TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
                
                'Print the Total Stock Volume in the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = TotalStockVol
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset the Total Stock Volume
                TotalStockVol = 0
            
            Else
            
                TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
            
            End If
        
        Next i
        
        'Identify last row in summary table
        LastRowSummaryTable = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To LastRowSummaryTable
        
            'Color Conditional Format for Yearly Change
            If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
            
            'Calculate & Print greatest % increase, greatest % decrease, greatest total volume and their ticker names
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRowSummaryTable)) Then
            ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).NumberFormat = "0.00%"
            
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRowSummaryTable)) Then
            ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).NumberFormat = "0.00%"
                    
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRowSummaryTable)) Then
            ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            
            End If
            
            LastRowSummaryTable = LastRowSummaryTable + 1
            
            
        Next i
            

  Next ws


End Sub
