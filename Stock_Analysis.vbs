Sub StockMarket()

Dim ticker As String
Dim Ws As Worksheet
Dim lastRow As Single
Dim Summary_Table_Index As Single

For Each Ws In Worksheets
    Summary_Table_Index = 2
    Ws.Cells(1, "I").Value = "Ticker"
    lastRow = Ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    For RowIndex = 2 To lastRow
        If Ws.Cells(RowIndex + 1, 1).Value <> Ws.Cells(RowIndex, 1).Value Then
            Ws.Cells(Summary_Table_Index, "I").Value = Ws.Cells(RowIndex, "A").Value
            Summary_Table_Index = Summary_Table_Index + 1
        End If
    Next RowIndex
Next Ws
End Sub

