Sub Bonus_Stock_Market_Analysis()
  Dim i As Long
  Dim last_row_count As Long
  Dim ws As Worksheet
  

Range("p1").Value = "Ticker"
Range("q1").Value = "Value"
Range("o2").Value = "Greatest % Increase"
Range("o3").Value = "Greatest % Decrease"
Range("o4").Value = "Greatest Total Volume"

Range("q2").Value = Cells(2, 11)
Range("q3").Value = Cells(2, 11)
Range("q4").Value = Cells(2, 12)

For Each ws In Worksheets
   last_row_count = ws.Cells(Rows.Count, "I").End(xlUp).Row
     For i = 2 To last_row_count
         If ws.Cells(i, 11).Value > Range("q2").Value Then
           Range("q2").Value = "%" & ((ws.Cells(i, 11).Value) * 100)
           Range("p2").Value = ws.Cells(i, 9).Value
         ElseIf ws.Cells(i, 11).Value < Range("q3").Value Then
                Range("q3").Value = "%" & ((ws.Cells(i, 11).Value) * 100)
                Range("p3").Value = ws.Cells(i, 9).Value
         ElseIf ws.Cells(i, 12).Value > Range("q4").Value Then
                  Range("q4").Value = ws.Cells(i, 12).Value
                  Range("p4").Value = ws.Cells(i, 9).Value
         End If
     Next i

Next ws

End Sub



