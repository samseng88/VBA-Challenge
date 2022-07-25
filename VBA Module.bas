Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data()

'Define variables
    Dim ticker As String
    Dim vol As Double
    Dim stock_data As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim ws As Worksheet

    For Each ws In Worksheets
    
            vol = 0
    
        'Title Headers
        ws.Cells(1, 9).Value = "ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Yearly Percentage"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        stock_data = 2
        
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
            
        'loops
            If year_open = 0 Then
                    year_open = ws.Cells(i, 3).Value
             
            ElseIf ws.Cells(i - 1, 1) = ws.Cells(i, 1) And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                year_close = ws.Cells(i, 6).Value
                yearly_change = year_close - year_open
                year_percentage = Format(yearly_change / year_open, "0.00%")
                ticker = ws.Cells(i, 1).Value
                vol = vol + ws.Cells(i, 7).Value
                
            'Value for each unique record
                ws.Cells(stock_data, 9).Value = ticker
                ws.Cells(stock_data, 10).Value = yearly_change
                ws.Cells(stock_data, 11).Value = year_percentage
                ws.Cells(stock_data, 12).Value = vol
                
              'Need to format cells
                If ws.Cells(stock_data, 10).Value <= 0 Then
                    ws.Cells(stock_data, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(stock_data, 10).Value > 0 Then
                    ws.Cells(stock_data, 10).Interior.ColorIndex = 4
                End If
                
                stock_data = stock_data + 1
                    vol = 0
            
            Else
                    vol = vol + ws.Cells(i, 7).Value
            End If
            
            Next i
            
        Next ws
        
    
    End Sub
    
    
