Attribute VB_Name = "Module1"
Sub StockData()

For Each ws In Worksheets

' Set an initial variable for holding the Ticker
  Dim Ticker As String

  ' Set an initial variable for Volume Total per Ticker
  Dim VolumeTotal As Double
  VolumeTotal = 0
  
  ' Set Up Minimum Date
  Dim OpenPrice As Double
  
  ' Set Up Max Close Date
  Dim ClosePrice As Double
  
  ' Set Up YearlyChange
  
  Dim YearlyChange As Double
  
  ' Set Up PercentChange
  
  Dim PercentChange As Single
  
  ' Keep track of the location for each Ticker
  Dim TickerSummaryRow As Integer
  TickerSummaryRow = 2

  Dim lastRow As Long
  lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Title Headers
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"

  ' Set Initial Open Price
  
  OpenPrice = ws.Cells(2, 3).Value
    

  ' Loop through all Tickers
  For i = 2 To lastRow


    ' Check if if the next row has the same name as current ticker
    ' if it is...
    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then

      ' Add to the VolumeTotal
      VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value

    ' if it's not (we're on the last row for this brand)...
    Else

      ' Set the Ticker name
      Ticker = ws.Cells(i, 1).Value
      
      ' Add to the VolumeTotal
      VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
      
      ' Set DateMaxClose
      
      ClosePrice = ws.Cells(i, 6).Value
           
      
      'Yearly Change
      
      YearlyChange = ClosePrice - OpenPrice
      
      'Percent Change
      
        
    If (OpenPrice = 0 And ClosePrice = 0) Then
        
        PercentChange = 0
        
        ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
        
        PercentChange = 1
        
        
        Else
        
        PercentChange = YearlyChange / OpenPrice
      
     End If
     
      ' Print this tickers's info in the summary table
      ws.Cells(TickerSummaryRow, 9).Value = Ticker
      ws.Cells(TickerSummaryRow, 10).Value = YearlyChange
      ws.Cells(TickerSummaryRow, 11).Value = PercentChange
      ws.Cells(TickerSummaryRow, 11).NumberFormat = "0.00%"
      ws.Cells(TickerSummaryRow, 12).Value = VolumeTotal

      ' Add one to the summary table row
      TickerSummaryRow = TickerSummaryRow + 1

      ' Reset the VolumeTotal
      VolumeTotal = 0
      
      ' Reset OpenPrice
      
      OpenPrice = ws.Cells(i + 1, 3).Value

    ' If the cell immediately following a row is the same brand...


    End If

  Next i
  
   ' Determine the Last Row of Yearly Change per WS
        YCLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        ' Set the Cell Colors
        For j = 2 To YCLastRow
            If (ws.Cells(j, 11).Value > 0 Or ws.Cells(j, 11).Value = 0) Then
                ws.Cells(j, 11).Interior.ColorIndex = 10
            ElseIf ws.Cells(j, 11).Value < 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 3
            End If
        Next j
Next ws

End Sub


