Sub VBA_Stocks()
    Dim ticker_symbol As String
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim total_Stock_Volume As Double
    Dim summary_table_row As Integer
    Dim Header_row As Integer
    Dim last_Row As Long
    Dim Hoja As Worksheet
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Variant

    
For Each Hoja In ActiveWorkbook.Sheets
    Hoja.Activate
    
    summary_table_row = 2
    Header_row = 1
    last_Row = Range("A" & Rows.Count).End(xlUp).Row
    
    Cells(Header_row, 9) = "Ticker"
    Cells(Header_row, 10) = "Year Change"
    Cells(Header_row, 11) = "Percentage Change"
    Cells(Header_row, 12) = "Total Stock Volume"
        
    
    Cells(Header_row, 15) = ""
    Cells(Header_row, 16) = "Ticker"
    Cells(Header_row, 17) = "Value"
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Drecrease"
    Cells(4, 15) = "Greatest Total Volume"
    
    
    
    For y = 2 To last_Row
        
        If Cells(y, 1).Value <> Cells(y + 1, 1).Value Then
        

            ticker_symbol = Cells(y, 1).Value
            total_volume = total_volume + Cells(y, 7).Value
            yearly_change = Cells(y - 1, 6).Value - Cells(y + 1, 3).Value
            
                If Cells(y + 1, 3).Value = 0 Then
                percentage_change = 0
                Else
                percentage_change = ((Cells(y - 1, 6).Value / (Cells(y + 1, 3).Value)) - 1)
                End If
            
            Cells(summary_table_row, 9).Value = ticker_symbol
            Cells(summary_table_row, 10).Value = yearly_change
            Cells(summary_table_row, 11).Value = Format(percentage_change, "0.00%")
            Cells(summary_table_row, 12).Value = total_volume
           
            
            If yearly_change > 0 Then
            Cells(summary_table_row, 10).Interior.ColorIndex = 4
            Else
            Cells(summary_table_row, 10).Interior.ColorIndex = 3
            End If
            
            summary_table_row = summary_table_row + 1
            total_volume = 0
            
        Else
        
            total_volume = total_volume + Cells(y, 7).Value
                 
        End If
    Next y
    
        greatest_increase = Application.WorksheetFunction.Max(Range("K:K"))
            Cells(2, 17).Value = Format(greatest_increase, "0.00%")
        greatest_decrease = Application.WorksheetFunction.Min(Range("K:K"))
            Cells(3, 17).Value = Format(greatest_decrease, "0.00%")
        greatest_volume = Application.WorksheetFunction.Max(Range("L:L"))
            Cells(4, 17).Value = greatest_volume
         greatest_volume = CDec(greatest_volume)
    
    lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
        For o = 2 To lastrow2
            If Cells(o, 11).Value = greatest_increase Then
                Cells(2, 16).Value = Cells(o, 9).Value
            ElseIf Cells(o, 11).Value = greatest_decrease Then
                Cells(3, 16).Value = Cells(o, 9).Value
            ElseIf Cells(o, 12).Value = greatest_volume Then
                Cells(4, 16).Value = Cells(o, 9).Value
            End If
        Next o
     
    
Next Hoja
   
   
End Sub