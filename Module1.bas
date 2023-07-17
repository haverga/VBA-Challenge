Attribute VB_Name = "Module1"
Sub Challenge2()
    Dim i As Long
    Dim r As Long
    Dim sum As Double
    Dim Change As Double
    Dim percentage As Double
    Dim ws As Worksheet
    Dim last_row As Long
    Dim j As Integer
    
    
    
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    r = 2
    j = 0
    sum = 0
    
    
    For i = 2 To last_row
        sum = sum + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(r, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(r, 12).Value = sum
            sum = 0
            Change = (ws.Cells(i, 6) - ws.Cells(r, 3))
            percentage = Change / ws.Cells(r, 3)
            ws.Cells(r, 10).Value = Change
            ws.Cells(r, 11).Value = percentage
            r = r + 1
            Range("K" & 2 + j).NumberFormat = "0.00%"
            Change = 0
            j = j + 1
        End If
            
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
                Else
                ws.Cells(i, 10).Value = ""
            End If
     
     Next i
     
     ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & last_row)) * 100
     ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & last_row)) * 100
     ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & last_row))
     
     
     increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & last_row)), ws.Range("K2:K" & last_row), 0)
        ws.Range("P2") = ws.Cells(increase_number + 1, 9).Value
     decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & last_row)), ws.Range("K2:K" & last_row), 0)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9).Value
     volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & last_row)), ws.Range("L2:L" & last_row), 0)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9).Value
    
    
    
    Next ws
    
End Sub
