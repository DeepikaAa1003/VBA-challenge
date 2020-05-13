
Sub ConsolidateDataFinal()

' loop through each worksheet
' Select the worksheet
' set the header of cells I1 to Ticker, J1 to Yearly Change, K1 to percentage change and L1 to Total Stock Volumne
' set rowcounter to 2 to write data
' set OpenYearAmount = Cells(2,3)
' calculate lastrow and last column of current worksheet
' loop from 2nd row to last row -1
' calculate TotalStockVolume = TotalStokVolume + Cells(i,7)
' If cells value not equal to next cell value then
' set ClosingYearAmount = Cells(i,6)
' calculate YearlyChange = ClosingYearAmount - OpenYearAmount 
' Calculate percentageChange =YearlyChange / OpenYearAmount
' Set cells(rowcounter,9) = cells(i,1) to set ticker name
' Set cells(rowcounter,10) = YearlyChange set format to decimal amount
' Set cells(rowcounter,11) = percentageChange set format to percentage
' Set cells(rowcounter,12) = TotalStockVolume
' Set TotalStockVolume = 0
' set rowounter = rowcounter + 1
' Calculate stockvolume, yearly change and precentage chaange for last ticker at the end
' Calculate Max and min of percentage change and mx for total stock volume to figure out the challenge part


Dim rowCounter As Integer
Dim OpenYearAmount As Double
Dim TotalStockVolume
Dim ClosingYearAmount As Double
Dim YearlyChange As Double
Dim percentageChange As Double

   'Loop through each worksheet

    For Each ws In Worksheets
        
        'Locate last column and last row of active worksheet
        
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        ' Set heading for cells in current worksheet
        ws.Cells(1, LastColumn + 2) = "Ticker"
        ws.Cells(1, LastColumn + 3) = "Yearly Change"
        ws.Cells(1, LastColumn + 4) = "Percentage Change"
        ws.Cells(1, LastColumn + 5) = "Total Stock Volume"
        
        ws.Cells(1, LastColumn + 9) = "Ticker"
        ws.Cells(1, LastColumn + 10) = "Value"
        ws.Cells(2, LastColumn + 8) = "Greatest % Increase"
        ws.Cells(3, LastColumn + 8) = "Greatest % Decrease"
        ws.Cells(4, LastColumn + 8) = "Greatest Total Volume"
        
        ' Set open year amount for first ticker
        OpenYearAmount = ws.Cells(2, 3)
        
        'Set row counter value to write data
        rowCounter = 2
                    
                    'loop through all rows of current worksheet
                    For i = 2 To LastRow - 1
                        
                        'Add stockvolume for each row of same ticker
                        TotalStockVolume = TotalStockVolume + ws.Cells(i, LastColumn)
                        
                        'Check if ticker value of current cell same as next cell
                        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                            
                            'Set closingYearAmount of ticker
                            ClosingYearAmount = ws.Cells(i, LastColumn - 1)
                            
                            'Calculate yearlyChange and percentageChange
                            YearlyChange = ClosingYearAmount - OpenYearAmount
                            If (OpenYearAmount > 0) Then
                                    percentageChange = (YearlyChange / OpenYearAmount)
                            Else
                                    
                                    percentageChange = 0
                        
                            End If
                            'Set Open year amount for next ticker
                            
                            OpenYearAmount = ws.Cells(i + 1, 3)
                            
                            'Write data of ticker to row denoted by rowcounter value
                            
                            ws.Cells(rowCounter, LastColumn + 2) = ws.Cells(i, 1)
                            ws.Cells(rowCounter, LastColumn + 3) = YearlyChange
                            ws.Cells(rowCounter, LastColumn + 4) = percentageChange
                            ws.Cells(rowCounter, LastColumn + 5) = TotalStockVolume
                            
                            'Apply number formatting for yearly change and percentage change
                            
                            ws.Cells(rowCounter, LastColumn + 3).NumberFormat = "0.00"
                            ws.Cells(rowCounter, LastColumn + 4).NumberFormat = "0.00%"
                            
                            'Apply color formatting of red and green
                            
                            If (ws.Cells(rowCounter, LastColumn + 3) > 0) Then
                                ws.Cells(rowCounter, LastColumn + 3).Interior.ColorIndex = 4
                            Else
                                ws.Cells(rowCounter, LastColumn + 3).Interior.ColorIndex = 3
                            End If
                            
                            'Reset total stock volumne and increment rowcounter value to write data for next ticker
                            
                            TotalStockVolume = 0
                            rowCounter = rowCounter + 1
                            
                        End If
                        
                    Next i
                    
                    
                    'Calculate TotalStockVolume, YearlyChange and PercentageChange for last ticker
                    
                    TotalStockVolume = TotalStockVolume + ws.Cells(LastRow, LastColumn)
                    ClosingYearAmount = ws.Cells(LastRow, LastColumn - 1)
                    YearlyChange = ClosingYearAmount - OpenYearAmount
                            If (OpenYearAmount > 0) Then
                                    percentageChange = (YearlyChange / OpenYearAmount)
                            Else
                                    
                                    percentageChange = 0
                        
                            End If
                    
                    ws.Cells(rowCounter, LastColumn + 2) = ws.Cells(LastRow, 1)
                    ws.Cells(rowCounter, LastColumn + 3) = YearlyChange
                    ws.Cells(rowCounter, LastColumn + 4) = percentageChange
                    ws.Cells(rowCounter, LastColumn + 5) = TotalStockVolume
                    
                    'Apply number formatting for yearly change and percentage change
                     
                    ws.Cells(rowCounter, LastColumn + 3).NumberFormat = "0.00"
                    ws.Cells(rowCounter, LastColumn + 4).NumberFormat = "0.00%"
                    
                    'Apply color formatting of red and green
                    
                    If (ws.Cells(rowCounter, LastColumn + 3) > 0) Then
                        ws.Cells(rowCounter, LastColumn + 3).Interior.ColorIndex = 4
                    Else
                        ws.Cells(rowCounter, LastColumn + 3).Interior.ColorIndex = 3
                    End If
                    
                    
                    'Calculate Greatest% Decrease and write the data
                                     
                    ws.Cells(2, LastColumn + 9) = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), 0), LastColumn + 2)
                    ws.Cells(2, LastColumn + 10) = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), 0), LastColumn + 4)
                    ws.Cells(2, LastColumn + 10).NumberFormat = "0.00%"
                    
                    'Calculate Greatest% Increase and write the data
                    
                    ws.Cells(3, LastColumn + 9) = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), 0), LastColumn + 2)
                    ws.Cells(3, LastColumn + 10) = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), 0), LastColumn + 4)
                    ws.Cells(3, LastColumn + 10).NumberFormat = "0.00%"
                                        
                    'Calculate Greatest Total Volume and write the data
                    
                    ws.Cells(4, LastColumn + 9) = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0), LastColumn + 2)
                    ws.Cells(4, LastColumn + 10).NumberFormat = "General"
                    ws.Cells(4, LastColumn + 10) = ws.Cells(WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0), LastColumn + 5)
                    
                    
                    ws.Columns("A:Q").AutoFit
                    
                    
           Next ws


End Sub



