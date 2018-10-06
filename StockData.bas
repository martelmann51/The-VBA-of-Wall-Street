Attribute VB_Name = "Module1"
Sub StockData()

For Each ws In Worksheets

' Set an initial variable for holding the Ticker
  Dim Ticker As String

  ' Set an initial variable for Volume Total per Ticker
  Dim VolumeTotal As Double
  VolumeTotal = 0

  
  ' Keep track of the location for each Ticker
  Dim TickerSummaryRow As Integer
  TickerSummaryRow = 2

  Dim lastRow As Long
  lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Title Headers
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Total Stock Volume"
    
  ' Loop through all Tickers
  For i = 2 To lastRow


    ' Check if if the next row has the same name as current ticker
    ' if it is...
    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then

      ' Add to the VolumeTotal
      VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value

    ' if it's not (we're on the last row for this ticker)...
    Else

      ' Set the Ticker name
      Ticker = ws.Cells(i, 1).Value
      
      ' Add to the VolumeTotal
      VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
           
      
     
      ' Print this tickers's info in the summary table
      ws.Cells(TickerSummaryRow, 9).Value = Ticker
      ws.Cells(TickerSummaryRow, 10).Value = VolumeTotal

      ' Add one to the summary table row
      TickerSummaryRow = TickerSummaryRow + 1

      ' Reset the VolumeTotal
      VolumeTotal = 0
      


    End If

  Next i
  
Next ws

End Sub


