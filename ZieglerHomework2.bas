Attribute VB_Name = "Module1"
Sub VBA_Of_Wall_Street()
For Each ws In Worksheets
    ' 1. Add titles to columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' 2. Put in the ticker
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    Dim WritRow As Long
    WritRow = 2
    Dim Open_Price As Double
    Open_Price = 0
    Dim Closing_Price As Double
    Closing_Price = 0
    Dim Price_Change As Double
    Price_Change = 0
    Dim Total_Stock_Volume As Double
    
    For i = 2 To LastRow
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(WritRow, 9).Value = ws.Cells(i, 1).Value
            WritRow = WritRow + 1
        End If
        
    ' 3. Populate the Yearly Change
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            Open_Price = ws.Cells(i, 3).Value
            
        End If
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Closing_Price = ws.Cells(i, 6)
            Price_Change = (Closing_Price - Open_Price)
            ws.Cells(WritRow - 1, 10).Value = Price_Change
            
            ' 4. Populate the Percent Change
            Dim Percent_Change As Double
            Percent_Change = 0
            If Open_Price <> 0 Then
                Percent_Change = (Price_Change / Open_Price) * 100
            End If
            ws.Cells(WritRow - 1, 11).Value = "%" & Round(Percent_Change, 2)
            If Price_Change > 0 Then
                ws.Cells(WritRow - 1, 10).Interior.ColorIndex = 4
            ElseIf Price_Change < 0 Then
                ws.Cells(WritRow - 1, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(WritRow - 1, 10).Interior.ColorIndex = 0
            End If
            
            ' 5. Populate the Total Stock Value
            ws.Cells(WritRow - 1, 12).Value = Total_Stock_Volume
            Total_Stock_Volume = 0

        End If
          
    Next i
    
Next ws
    
End Sub

