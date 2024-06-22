Attribute VB_Name = "Module1"
Sub Stock_Data()

For Each ws In Worksheets
WorksheetName = ws.Name

ws.Cells(1, 10).value = "Ticker"
ws.Cells(1, 11).value = "Quartley Change"
ws.Cells(1, 12).value = "Percent Change"
ws.Cells(1, 13).value = "Total Stock Volume"

    

Dim Ticker_Symbol As String
Dim Quarterly_Change As Single
Dim Percent_Change As Single
Dim Total_Stock_Volume As Variant
Dim Summary_Table_Row As Variant
Dim Open_Price As Variant
Dim Close_Price As Double
Dim firstRow As Variant






Quarterly_Change = 0
Percent_Change = 0
Total_Stock_Volume = 0
Summary_Table_Row = 2
Open_Price = 0
Close_Price = 0

For i = 2 To 93001
    
    If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
    
        Ticker_Symbol = ws.Cells(i, 1).value
        
        Close_Price = ws.Cells(i, 6).value
        
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).value
        

        
        
        ws.Range("J" & Summary_Table_Row).value = Ticker_Symbol
        
        ws.Range("K" & Summary_Table_Row).value = Quarterly_Change
        
        ws.Range("L" & Summary_Table_Row).value = Percent_Change
        
        ws.Range("M" & Summary_Table_Row).value = Total_Stock_Volume
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        
        
        Percent_Change = 0
        Quarterly_Change = 0
        Open_Price = 0
        Close_Price = 0
        Total_Stock_Volume = 0
        
    Else
    
        Quarterly_Change = Open_Price - Close_Price
        
        Percent_Change = (ws.Cells(i, 6).value - ws.Cells(i, 3).value) / ws.Cells(i, 3).value
        
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).value
        
        
        End If
        
            For j = 3 To 3
               If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
    
        Open_Price = ws.Cells(i, 3).value
        
                
            
                End If
                
                Next j
            Next i
                
    
          
  
Next ws

    
End Sub


Sub volumes()

For Each ws In Worksheets
WorksheetName = ws.Name

ws.Cells(1, 15).value = "Ticker"
ws.Cells(1, 16).value = "Greatest % Increase"
ws.Cells(1, 17).value = "Greatest % Decrease"
ws.Cells(1, 18).value = "Greatest Total Volume"

Dim Ticker_Symbol As String
Dim Total_Stock_Volume As Variant
Dim Greatest_Increase As Variant
Dim Greatest_Decrease As Variant
Dim Greatest_Total_volume As Variant
Dim Summary_Table_Row As Variant

Columns.AutoFit

Greatest_Increase = 0
Greatest_Decrease = 0
Greatest_Total_volume = 0
Summary_Table_Row = 2


    For i = 2 To 93001
    
        If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
    
        Ticker_Symbol = ws.Cells(i, 15).value
        
        Greatest_Increase = Application.WorksheetFunction.Max(Range("G:G"))
        
        Greatest_Decrease = Application.WorksheetFunction.Min(Range("G:G"))
  
        'Greatest_Total_volume = Greatest_Total_volume + ws.Cells(i, 7).value
        

        
        
        ws.Range("O" & Summary_Table_Row).value = Ticker_Symbol
        
        ws.Range("P" & Summary_Table_Row).value = Greatest_Decrease
        
        ws.Range("Q" & Summary_Table_Row).value = Greatest_Increase
        
        'ws.Range("R" & Summary_Table_Row).value = Total_Stock_Volume
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        
        
        Greatest_Decrease = 0
        Greatest_Decrease = 0
        Greatest_Total_volume = 0
        
    Else
         Ticker_Symbol = ws.Cells(i, 15).value
        
        Greatest_Increase = Application.WorksheetFunction.Max(Range("G:G"))
        
        Greatest_Decrease = Application.WorksheetFunction.Min(Range("G:G"))
  
        'Greatest_Total_volume = Greatest_Total_volume + ws.Cells(i, 7).value
        
        
        End If
Next i
Next ws

End Sub



