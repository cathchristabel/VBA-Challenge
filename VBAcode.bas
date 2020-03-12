Attribute VB_Name = "Module1"
Sub StockMarket()

    ' Set dimensions
    Dim Ticker As String
    Dim i, start, RowCount As Long
    Dim j As Integer
    Dim TotalStock, percentchange As Double

    For Each ws In Worksheets

        ' Set title row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Price Change"
        ws.Range("K1").Value = "Yearly Percent Change"
        ws.Range("L1").Value = "Total Stock Number"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
        ' get the row number of the last row with data
        RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set values for each worksheet
        j = 2
        TotalStock = 0
        start = 2

        For i = 2 To RowCount
        
            ' If ticker changes then print results
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                ' Stores results in variables
                Ticker = ws.Cells(i, 1).Value
                TotalStock = TotalStock + ws.Cells(i, 7).Value
                
                ' Handle zero total volume
                If TotalStock = 0 Then
                    ' print the results
                    ws.Range("I" & j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & j).Value = 0
                    ws.Range("K" & j).Value = "%" & 0
                    ws.Range("L" & j).Value = 0
    
                Else
                    ' Find First non zero starting value
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                Exit For
                            End If
                         Next find_value
                    End If
                
                    pricechange = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percentchange = Round((pricechange / ws.Cells(start, 3) * 100), 4)
                    
                    ' start of the next stock ticker
                    start = i + 1
                    
                    ' Print the results to specific cells
                    ws.Cells(j, 9).Value = Ticker
                    ws.Cells(j, 10).Value = pricechange
                    ws.Cells(j, 11).Value = "%" & percentchange
                    ws.Cells(j, 12).Value = TotalStock
                
                    ' Colors positives green and negatives red
                    If ws.Cells(j, 10).Value > 0 Then
                        ws.Cells(j, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(j, 10).Value < 0 Then
                        ws.Cells(j, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(j, 10).Interior.ColorIndex = 0
                    End If
                    
                End If
       
                ' Reset variables for new stock ticker
                TotalStock = 0
                j = j + 1


            ' If ticker is still the same, adds stock volume
            Else
                TotalStock = TotalStock + ws.Cells(i, 7).Value
            End If
        
        Next i
        
        ' take the max and min and place them in a separate part in the worksheet
        ws.Cells(2, 17).Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
        ws.Cells(3, 17).Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
        ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
        
        ' returns one less because header row not a factor
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)

        ' final ticker symbol for  total, greatest % of increase and decrease, and average
        ws.Cells(2, 16).Value = ws.Cells(increase_number + 1, 9)
        ws.Cells(3, 16).Value = ws.Cells(decrease_number + 1, 9)
        ws.Cells(4, 16).Value = ws.Cells(volume_number + 1, 9)

        ws.Columns("I:Q").AutoFit
                
    Next ws

End Sub
