Sub testing()
    

    'Loop through all sheet
    For Each ws In Worksheets
        
        'Determine last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Naming new Columns
        Dim TC As Integer
        TC = 9
        ws.Cells(1, TC).Value = "Ticker"
        ws.Cells(1, TC + 1).Value = "Yearly Change"
        ws.Cells(1, TC + 2).Value = "Percent Change"
        ws.Cells(1, TC + 3).Value = "Total stock volume"
        
        Dim Open_price As Double
        Open_price = ws.Cells(2, 3).Value
        Dim Close_price As Double
        Dim Row As Integer
        Row = 2
        volume = ws.Cells(2, 7).Value
        
        For i = 2 To lastrow
        
            'Find Total stock Volume
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                volume = volume + ws.Cells(i + 1, 7).Value
            
            Else
                ws.Cells(Row, TC + 3) = volume
                volume = 0
                
                'Find Ticker Name
                ws.Cells(Row, TC).Value = ws.Cells(i, 1).Value
              
              'Find Yearly Change
                Close_price = ws.Cells(i, 6).Value
                Yearly_change = Close_price - Open_price
                ws.Cells(Row, TC + 1).Value = Yearly_change
                
                'Change Yearly Change Column color
                If ws.Cells(Row, TC + 1).Value > 0 Then
                    ws.Cells(Row, TC + 1).Interior.Color = vbGreen
                ElseIf ws.Cells(Row, TC + 1).Value < 0 Then
                    ws.Cells(Row, TC + 1).Interior.Color = vbRed
                End If

                'Find Percent change
                If Open_price = 0 And Yearly_change = 0 Then
                    Percent = 0
                ElseIf Open_price = 0 And Yearly_change <> 0 Then
                    Percent = 1
                Else
                    Percent = Yearly_change / Open_price
                End If

                ws.Cells(Row, TC + 2).Value = Percent
                ws.Cells(Row, TC + 2).NumberFormat = "0.00%"
                
                'Reset open price for nexr ticker
                Open_price = ws.Cells(i + 1, 3).Value
                Row = Row + 1
    
             End If
        Next i
        
        'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
        
        ws.Cells(2, TC + 5).Value = "Greatest % Increase"
        ws.Cells(3, TC + 5).Value = "Greatest % Decrease"
        ws.Cells(4, TC + 5).Value = "Greatest Total Volume"
        ws.Cells(1, TC + 6).Value = "Ticker"
        ws.Cells(1, TC + 7).Value = "Value"
        Min = 0
        Max = 0
        TMax = 0
        For j = 2 To Row
            If ws.Cells(j, TC + 2).Value < Min Then
            Min = ws.Cells(j, TC + 2).Value
            ws.Cells(3, TC + 7).Value = ws.Cells(j, TC + 2).Value
            ws.Cells(3, TC + 6).Value = ws.Cells(j, TC).Value
            End If

            If ws.Cells(j, TC + 2).Value > Max Then
                Max = ws.Cells(j, TC + 2).Value
                ws.Cells(2, TC + 7).Value = ws.Cells(j, TC + 2).Value
                ws.Cells(2, TC + 6).Value = ws.Cells(j, TC).Value
            End If

            If ws.Cells(j, TC + 3).Value > TMax Then
                TMax = ws.Cells(j, TC + 3).Value
                ws.Cells(4, TC + 7).Value = ws.Cells(j, TC + 3).Value
                ws.Cells(4, TC + 6).Value = ws.Cells(j, TC).Value
            End If
        Next j
        ws.Cells(2, TC + 7).NumberFormat = "0.00%"
        ws.Cells(3, TC + 7).NumberFormat = "0.00%"
        
    Next ws
End Sub

