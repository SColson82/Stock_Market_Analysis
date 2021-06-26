Sub StockMarket()

Dim Ticker As String
Dim TickerOpenPrice As Double
Dim TickerClosePrice As Double
Dim TotalStockVolume As Double
Dim SummaryTableIndex As Double
Dim PercentIncrease As Double
Dim PercentIncreaseValue As Double
Dim PercentDecrease As Double
Dim GreatestTotalVolume As Double

For Each Ws In Worksheets

    TotalStockVolume = 0
    SummaryTableIndex = 2
    PercentIncrease = 0
    PercentDecrease = 0
    GreatestTotalVolume = 0
    TickerOpenPrice = Ws.Cells(2, "C").Value
    
    Ws.Cells(1, "I").Value = "Ticker"
    Ws.Cells(1, "J").Value = "YearlyChange"
    Ws.Cells(1, "K").Value = "Percentage Change"
    Ws.Cells(1, "L").Value = "Total Stock Volume"
    Ws.Cells(1, "P").Value = "Ticker"
    Ws.Cells(1, "Q").Value = "Value"
    Ws.Cells(2, "O").Value = "Greatest % Increase"
    Ws.Cells(3, "O").Value = "Greatest % Decrease"
    Ws.Cells(4, "O").Value = "Greatest Total Volume"
    
    
    For RowIndex = 2 To Ws.Cells(Rows.Count, 1).End(xlUp).Row
        TotalStockVolume = TotalStockVolume + Ws.Cells(RowIndex, 7).Value
        Ticker = Ws.Cells(RowIndex, 1).Value
            If Ws.Cells(RowIndex + 1, 1).Value <> Ws.Cells(RowIndex, 1).Value Then
                Ws.Cells(SummaryTableIndex, "I").Value = Ticker
                TickerClosePrice = Ws.Cells(RowIndex, "F").Value
                Ws.Cells(SummaryTableIndex, "J").Value = TickerClosePrice - TickerOpenPrice
 '               Ws.Cells(3, 16).Value =
  '              Ws.Cells(4, 16).Value =
                If TickerOpenPrice <> 0 Then
                    'Formatted to show Percentage Sign
                    Ws.Cells(SummaryTableIndex, "K").Value = FormatPercent((TickerClosePrice - TickerOpenPrice) / TickerOpenPrice, 2)
                Else
                    Ws.Cells(SummaryTableIndex, "K").Value = Null
                End If
                Ws.Cells(SummaryTableIndex, "L").Value = TotalStockVolume
                If Ws.Cells(SummaryTableIndex, "J").Value > 0 Then
                    Ws.Cells(SummaryTableIndex, "J").Interior.ColorIndex = 4
                Else
                    Ws.Cells(SummaryTableIndex, "J").Interior.ColorIndex = 3
                End If
                If Ws.Cells(SummaryTableIndex, "K").Value > PercentIncrease Then
                    PercentIncrease = Ws.Cells(SummaryTableIndex, "K").Value
                    PercentIncreaseTotal = RowIndex
                End If

                TickerOpenPrice = Ws.Cells(RowIndex + 1, "C").Value
                TotalStockVolume = 0
                SummaryTableIndex = SummaryTableIndex + 1
            End If
    Next RowIndex
    
    Ws.Cells(2, "P").Value = PercentIncreaseTotal
    Ws.Cells(2, "Q").Value = PercentIncrease
    Ws.Columns("A:Q").AutoFit
Next Ws
End Sub