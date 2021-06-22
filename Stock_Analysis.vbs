Sub StockMarket()

Dim ticker As String
Dim Ws As Worksheet
Dim lastRow As Double
Dim Summary_Table_Index As Double
Dim TickerOpenPrice As Double
Dim TickerClosePrice As Double

For Each Ws In Worksheets
    Summary_Table_Index = 2
    TickerOpenPrice = Ws.Cells(2, "C").Value
    Ws.Cells(1, "I").Value = "Ticker"
    Ws.Cells(1, "J").Value = "YearlyChange"
    Ws.Cells(1, "K").Value = "Percentage Change"
    lastRow = Ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    For RowIndex = 2 To lastRow
        If Ws.Cells(RowIndex + 1, 1).Value <> Ws.Cells(RowIndex, 1).Value Then
            Ws.Cells(Summary_Table_Index, "I").Value = Ws.Cells(RowIndex, "A").Value
            TickerClosePrice = Ws.Cells(RowIndex, "F").Value
            Ws.Cells(Summary_Table_Index, "J").Value = TickerClosePrice - TickerOpenPrice
            If TickerOpenPrice <> 0 Then
                Ws.Cells(Summary_Table_Index, "K").Value = Round(((TickerClosePrice - TickerOpenPrice) * 100) / TickerOpenPrice, 2)
            Else
                Ws.Cells(Summary_Table_Index, "K").Value = 0
            End If
            TickerOpenPrice = Ws.Cells(RowIndex + 1, "C").Value
            Summary_Table_Index = Summary_Table_Index + 1
        End If
    Next RowIndex
Next Ws
End Sub