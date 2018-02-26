Attribute VB_Name = "Module1"
Sub alpha_test():
   Dim last_row As Double
   Dim current_ticker As String
   Dim current_volume As Double
   Dim total_vol As Double
   Dim int_rows As Long
   Dim j_cols As Long
   
   j = 2
   total_vol = 0
   
   Cells(1, 9).Value = "Ticker"
   Cells(1, 10).Value = "Total Volume"
   
   last_row = Cells(Rows.Count, 1).End(xlUp).Row
   
   For i = 2 To last_row + 1
   
   current_ticker = Cells(i, 1).Value
   current_volume = Cells(i, 7).Value
   
   If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    current_ticker = Cells(i + 1, 1).Value
   Else
    current_ticker = Cells(i, 1).Value
    total_vol = total_vol + Cells(i, 7).Value
    Cells(j, 9).Value = current_ticker
    Cells(j, 10).Value = total_vol
    total_vol = 0
    j = j + 1
    
    End If
    Next i
    
    
   
   
   
End Sub

