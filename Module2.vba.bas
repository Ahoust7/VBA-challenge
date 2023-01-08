Attribute VB_Name = "Module1"

Sub Alphabetic_Testing()
For Each ws In Worksheets
Dim WorksheetName As String
Dim ticker As String
Dim total_stock_value As String
total_stock_value = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim percent_change As String
Dim yearly_change As Variant
Dim i As Long
Dim j As Long
WorksheetName = ws.Name

 ws.Cells(1, 9) = "Ticker"
 ws.Cells(1, 10) = "Yearly Change"
 ws.Cells(1, 11) = "Percent Change"
 ws.Cells(1, 12) = "Total Stock Volume"
 
 
 
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 
 For j = 2 To 2
 For i = 2 To lastrow

 
 
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 ticker = ws.Cells(i, 1).Value
 total_stock_value = total_stock_value + ws.Cells(i, 7).Value
 yearly_change = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
 percent_change = (ws.Cells(i, 6).Value / ws.Cells(j, 3).Value) - 1
     
 
 ws.Range("I" & Summary_Table_Row).Value = ticker
 ws.Range("J" & Summary_Table_Row).Value = yearly_change
 If ws.Cells(Summary_Table_Row, 10).Value < 0 Then
                
    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                
    Else
                
    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                
    End If
    
 ws.Range("K" & Summary_Table_Row).Value = percent_change
 ws.Cells(Summary_Table_Row, 11).Value = Format(percent_change, "Percent")
                    
     
 ws.Range("L" & Summary_Table_Row).Value = total_stock_value
 
 
 Summary_Table_Row = Summary_Table_Row + 1
 total_stock_value = 0
 j = i + 1
 Else
 total_stock_value = total_stock_value + Cells(i, 7).Value

 End If


Next i
Next j
Dim great_inc As Double
Dim great_decr As Double
Dim greatvol As Double



ws.Cells(1, 16) = "Ticker"
ws.Cells(1, 17) = "Value"
ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(4, 15) = "Greatest Total Volume"

great_inc = ws.Cells(2, 11).Value
great_dec = ws.Cells(2, 11).Value
greatvol = ws.Cells(2, 12).Value

LastrowI = Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To LastrowI

  If ws.Cells(i, 12).Value > greatvol Then
                greatvol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                greatvol = greatvol
                
                End If
                
If ws.Cells(i, 11).Value > great_inc Then
 great_inc = ws.Cells(i, 11).Value
 ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
 ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
 Else
           
 great_inc = great_inc
               
 End If

If ws.Cells(i, 11).Value < great_dec Then
 great_dec = ws.Cells(i, 11).Value
 ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
 ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
 Else
           
 great_dec = great_dec
               
 End If

 
 
ws.Cells(2, 17).Value = Format(great_inc, "Percent")
ws.Cells(3, 17).Value = Format(great_dec, "Percent")
ws.Cells(4, 17).Value = Format(greatvol, "Scientific")
    
Next i
Worksheets(WorksheetName).Columns("A:Z").AutoFit
Next ws
End Sub
