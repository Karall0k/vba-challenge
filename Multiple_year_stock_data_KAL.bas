Attribute VB_Name = "Module1"
Sub Stock_data()
 For Each ws In Worksheets
    
    
    'Set variables
    Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    Dim ticker_sym As String
    Dim Yr_change As Double
    Dim Percent_change As Double
    Dim Total_Stock_vol As Double
        Total_Stock_vol = 0
    Dim c As Integer
    Dim Sum_Table_Row As Integer
        Sum_Table_Row = 2
   
    Dim Open_price As Double
        Open_price = ws.Cells(2, 3).Value
    Dim Close_price As Double
    
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    
    'Loop Through Stocks
     For i = 2 To lastrow
      
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                 
           'Set Ticker symbol
            ticker_sym = ws.Cells(i, 1).Value
            'Store opening price
            Open_price = ws.Cells(i, 3).Value
             'Store closing price
            Close_price = ws.Cells(i, 6).Value
            'Add to Stock Volume
            Total_Stock_vol = Total_Stock_vol + ws.Cells(i, 7).Value
            'Print ticker; Symbol
            ws.Range("I" & Sum_Table_Row).Value = ticker_sym
            'Print Stock Volume
            ws.Range("L" & Sum_Table_Row).Value = Total_Stock_vol
           
           ' Calculate Yearly Change
           Yr_change = (Close_price - Open_price)
           'Print Yearly Change
           ws.Range("J" & Sum_Table_Row).Value = Yr_change
           ' Calculate Percent Change
                If Open_price = 0 Then
                     Percent_change = 0
                     
                 Else
                      Percent_change = Yr_change / Open_price
                End If
           'Print Percent Change
           ws.Range("K" & Sum_Table_Row).Value = Percent_change
           
           'Add row to summary table
           Sum_Table_Row = Sum_Table_Row + 1
           'Reset Stock Volume total and opening price
           Total_Stock_vol = 0
           Open_price = ws.Cells(i + 1, 3)
           
        'If Tickers in adjacent rows do match
        Else
           'Add to stock volume
           Total_Stock_vol = Total_Stock_vol + ws.Cells(i, 7).Value
                   
        End If
          
        Next i
   
 'Conditional Formatting
      For i = 2 To lastrow
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
       Next i
 

'Max_Min Summary()
       For i = 2 To lastrow
        
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
         Next i
    Next ws

End Sub
