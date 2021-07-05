Sub Stock_Market_Analysis()
  
  Dim i As Long
  Dim j As Integer
  Dim start_value As Long
  Dim years As Integer
  Dim total_stock As Double
  Dim find_value As Double
  Dim average_change As Double
  Dim percent_change As Double
  Dim change As Double
  Dim year_change As Double
  Dim last_row_count As Long
  Dim ws As Worksheet
  
For Each ws In Worksheets
j = 0
total_stock = 0
change = 0
year_change = 0
start_value = 2
   
ws.Range("i1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("l1").Value = "Total Stock Volume"

last_row_count = ws.Cells(Rows.Count, "A").End(xlUp).Row
  For i = 2 To last_row_count
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
        total_stock = total_stock + ws.Cells(i, 7).Value
        If total_stock = 0 Then
           ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
           ws.Range("j" & 2 + j).Value = 0
           ws.Range("k" & 2 + j).Value = "%" & 0
           ws.Range("l" & 2 + j).Value = 0
        Else
           If Cells(start_value, 3) = 0 Then
               For find_value = start_value To i
                   If ws.Cells(find_value, 3).Value <> 0 Then
                       start_value = find_value
                       Exit For
                   End If
               Next find_value
           End If
           If ws.Cells(start_value, 3) <> 0 Then
           change = (ws.Cells(i, 6) - ws.Cells(start_value, 3))
           percent_change = Round((change / ws.Cells(start_value, 3) * 100), 2)
           
           start_value = i + 1
           ws.Range("i" & 2 + j).Value = ws.Cells(i, 1).Value
           ws.Range("j" & 2 + j).Value = Round(change, 2)
           ws.Range("k" & 2 + j).Value = "%" & percent_change
           ws.Range("l" & 2 + j).Value = total_stock
           End If
           If change > 0 Then
              ws.Range("j" & 2 + j).Interior.ColorIndex = 4
           ElseIf change < 0 Then
              ws.Range("j" & 2 + j).Interior.ColorIndex = 3
           Else
              ws.Range("j" & 2 + j).Interior.ColorIndex = 0
           End If
         End If
     
     total_stock = 0
     change = 0
     j = j + 1
     years = 0
     year_change = 0
   Else
       total_stock = total_stock + ws.Cells(i, 7).Value
   End If
               

  Next i
Next ws

End Sub



