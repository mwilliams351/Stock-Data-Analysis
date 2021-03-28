Sub Stock_Fun()
    
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        
        lastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

    
        'Continue to check variable syntax
        Dim YearlyChange As Double
        Dim Ticker As String
        Dim PercentChange As Double
        Dim stock_volume As Double
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        stock_volume = 0
        Dim i As Long
        Dim open_price As Double
        Dim close_price As Double
        'Add titles
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        'This has to change for each summary row
       open_price = Cells(2, 3).Value
       
        
        For i = 2 To lastRow
         
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
              
                Ticker = Cells(i, 1).Value
                Cells(Row, 9).Value = Ticker
               
                close_price = Cells(i, 6).Value
               'Yearly change is last - first for each data set not the sum of every change
                YearlyChange = close_price - open_price
                Cells(Row, 10).Value = YearlyChange
                'fix division by zero errors
                If (open_price = 0 And close_price = 0) Then
                    PercentChange = 0
                ElseIf (open_price = 0 And close_price <> 0) Then
                    PercentChange = 1
                Else
                    PercentChange = YearlyChange / open_price
                    Cells(Row, 11).Value = PercentChange
                    Cells(Row, 11).NumberFormat = "0.00%"
                End If
                'Add same ticker symbols
                stock_volume = stock_volume + Cells(i, 7).Value
                Cells(Row, 12).Value = stock_volume
                'Remember to reset values
                Row = Row + 1
                
                open_price = Cells(i + 1, 3)
               
                stock_volume = 0
            
            Else
                stock_volume = stock_volume + Cells(i, 7).Value
            End If
        Next i
        
       
        YCLastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
       
        For j = 2 To YCLastRow
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

        
        For x = 2 To YCLastRow
            If Cells(x, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, 16).Value = Cells(x, 9).Value
                Cells(2, 17).Value = Cells(x, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(x, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, 16).Value = Cells(x, 9).Value
                Cells(3, 17).Value = Cells(x, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(x, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, 16).Value = Cells(x, 9).Value
                Cells(4, 17).Value = Cells(x, 12).Value
            End If
        Next x
        
    Next WS
        
End Sub
