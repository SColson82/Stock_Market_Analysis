VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockMarket()

Dim Ticker As String
Dim TickerOpenPrice, TickerClosePrice As Double
Dim TotalStockVolume As Double
Dim SummaryTableIndex As Double
Dim PercentIncrease, PercentDecrease, GreatestTotalVolumeTicker As String
Dim PercentIncreaseValue, PercentDecreaseValue, GreatestTotalVolume As Double

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
                If TickerOpenPrice <> 0 Then
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
                    PercentIncreaseTotal = Ws.Cells(SummaryTableIndex, "I").Value
                End If
                If Ws.Cells(SummaryTableIndex, "K").Value < PercentDecrease Then
                    PercentDecrease = Ws.Cells(SummaryTableIndex, "K").Value
                    PercentDecreaseTotal = Ws.Cells(SummaryTableIndex, "I").Value
                End If
                If Ws.Cells(SummaryTableIndex, "L").Value > GreatestTotalVolume Then
                    GreatestTotalVolume = Ws.Cells(SummaryTableIndex, "L").Value
                    GreatestTotalVolumeTicker = Ws.Cells(SummaryTableIndex, "I").Value
                End If
                TickerOpenPrice = Ws.Cells(RowIndex + 1, "C").Value
                TotalStockVolume = 0
                SummaryTableIndex = SummaryTableIndex + 1
            End If
    Next RowIndex
    Ws.Cells(2, "P").Value = PercentIncreaseTotal
    Ws.Cells(2, "Q").Value = (FormatPercent(PercentIncrease, 2))
    Ws.Cells(3, "P").Value = PercentDecreaseTotal
    Ws.Cells(3, "Q").Value = (FormatPercent(PercentDecrease, 2))
    Ws.Cells(4, "P").Value = GreatestTotalVolumeTicker
    Ws.Cells(4, "Q").Value = GreatestTotalVolume
    Ws.Columns("A:Q").AutoFit
Next Ws
End Sub





