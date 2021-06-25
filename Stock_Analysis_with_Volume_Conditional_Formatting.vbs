Sub StockMarket()

Dim Ticker As String
Dim TickerOpenPrice As Double
Dim TickerClosePrice As Double
Dim TotalStockVolume As Double
Dim SummaryTableIndex As Double

For Each Ws In Worksheets

    TotalStockVolume = 0
    SummaryTableIndex = 2
    TickerOpenPrice = Ws.Cells(2, "C").Value
    
    Ws.Cells(1, "I").Value = "Ticker"
    Ws.Cells(1, "J").Value = "YearlyChange"
    Ws.Cells(1, "K").Value = "Percentage Change"
    Ws.Cells(1, "L").Value = "Total Stock Volume"
    
    
    For RowIndex = 2 To Ws.Cells(Rows.Count, 1).End(xlUp).Row
        TotalStockVolume = TotalStockVolume + Ws.Cells(RowIndex, 7).Value
        Ticker = Ws.Cells(RowIndex, 1).Value
            If Ws.Cells(RowIndex + 1, 1).Value <> Ws.Cells(RowIndex, 1).Value Then
                Ws.Cells(SummaryTableIndex, "I").Value = Ticker
                TickerClosePrice = Ws.Cells(RowIndex, "F").Value
                Ws.Cells(SummaryTableIndex, "J").Value = TickerClosePrice - TickerOpenPrice
                If TickerOpenPrice <> 0 Then
                    Ws.Cells(SummaryTableIndex, "K").Value = Round(((TickerClosePrice - TickerOpenPrice) * 100) / TickerOpenPrice, 2)
                Else
                    Ws.Cells(SummaryTableIndex, "K").Value = Null
                End If
                Ws.Cells(SummaryTableIndex, "L").Value = TotalStockVolume
                If Ws.Cells(SummaryTableIndex, "J").Value > 0 Then
                    Ws.Cells(SummaryTableIndex, "J").Interior.ColorIndex = 4
                Else
                    Ws.Cells(SummaryTableIndex, "J").Interior.ColorIndex = 3
                End If

                TickerOpenPrice = Ws.Cells(RowIndex + 1, "C").Value
                TotalStockVolume = 0
                SummaryTableIndex = SummaryTableIndex + 1
        End If
    Next RowIndex
    Ws.Columns("A:L").AutoFit
Next Ws
End Sub