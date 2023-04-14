Attribute VB_Name = "Module1"
Sub GreatestIncrease()

Dim ws As Worksheet
Dim decTicker As String
Dim incTicker As String
Dim volTicker As String
Dim decrease As Double
decrease = 0
Dim increase As Double
increase = 0
Dim volume As Double
volume = 0


For Each ws In ThisWorkbook.Worksheets

  ws.Range("P" & 1).Value = "Ticker"
  ws.Range("Q" & 1).Value = "Value"
  
  ws.Range("O" & 2).Value = "Greatest % Increase"
  ws.Range("O" & 3).Value = "Greatest % Decrease"
  ws.Range("O" & 4).Value = "Greatest Total Volume"

For i = 2 To LastRow
    
        If ws.Cells(i, 11) < 0 Then
            If ws.Cells(i, 11) < decrease Then
            decrease = ws.Cells(i, 11).Value
            decTicker = ws.Cells(i, 9).Value
        End If
    
    
    If ws.Cells(i, 11) > 0 Then
        If ws.Cells(i, 11) > increase Then
            increase = ws.Cells(i, 11).Value
            incTicker = ws.Cells(i, 9).Value
    End If


    If ws.Cells(i, 12) > volume Then
        volume = ws.Cells(i, 12).Value
        volTicker = ws.Cells(i, 9).Value
    End If
    
ws.Range("P" & 2).Value = incTicker
ws.Range("P" & 3).Value = decTicker
ws.Range("P" & 4).Value = volTicker
ws.Range("Q" & 2).NumberFormat = "0.00%"
ws.Range("Q" & 2).Value = increase
ws.Range("Q" & 3).NumberFormat = "0.00%"
ws.Range("Q" & 3).Value = decrease
ws.Range("Q" & 4).Value = volume


ws.Range("P" & 2).Value = incTicker
ws.Range("P" & 3).Value = decTicker
ws.Range("P" & 4).Value = volTicker
ws.Range("Q" & 2).NumberFormat = "0.00%"
ws.Range("Q" & 2).Value = increase
ws.Range("Q" & 3).NumberFormat = "0.00%"
ws.Range("Q" & 3).Value = decrease
ws.Range("Q" & 4).Value = volume


  Next i
  
  Next ws
    
    

End Sub
