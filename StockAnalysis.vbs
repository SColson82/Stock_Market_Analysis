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

'Declare my variables
Dim Ticker As String
Dim TickerOpenPrice, TickerClosePrice As Double
Dim TotalStockVolume As Double
Dim SummaryTableIndex As Double
'Declare my variables for the Challenge
Dim PercentIncrease, PercentDecrease, GreatestTotalVolumeTicker As String
Dim PercentIncreaseValue, PercentDecreaseValue, GreatestTotalVolume As Double

'Apply to all Worksheets in the Workbook
For Each Ws In Worksheets

    'Set my initial conditions
    TotalStockVolume = 0
    SummaryTableIndex = 2
    PercentIncrease = 0
    PercentDecrease = 0
    GreatestTotalVolume = 0
    TickerOpenPrice = Ws.Cells(2, "C").Value
    
    'Apply Column and Row Titles for the data requested
    Ws.Cells(1, "I").Value = "Ticker"
    Ws.Cells(1, "J").Value = "YearlyChange"
    Ws.Cells(1, "K").Value = "Percentage Change"
    Ws.Cells(1, "L").Value = "Total Stock Volume"
    Ws.Cells(1, "P").Value = "Ticker"
    Ws.Cells(1, "Q").Value = "Value"
    Ws.Cells(2, "O").Value = "Greatest % Increase"
    Ws.Cells(3, "O").Value = "Greatest % Decrease"
    Ws.Cells(4, "O").Value = "Greatest Total Volume"
    
    'Set the RowIndex variable
    For RowIndex = 2 To Ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Define Total Stock Volume for each ticker
        TotalStockVolume = TotalStockVolume + Ws.Cells(RowIndex, 7).Value
        'Define Ticker
        Ticker = Ws.Cells(RowIndex, 1).Value
            'Divide the Data into sections dependent upon the ticker symbol to which it applies.
            If Ws.Cells(RowIndex + 1, 1).Value <> Ws.Cells(RowIndex, 1).Value Then
                'Present a list of the Tickers
                Ws.Cells(SummaryTableIndex, "I").Value = Ticker
                'Calculate the Close Price
                TickerClosePrice = Ws.Cells(RowIndex, "F").Value
                'Calculate the amount of change from open to close total.
                Ws.Cells(SummaryTableIndex, "J").Value = TickerClosePrice - TickerOpenPrice
                'Calculate the percent change, set any Null statement for stocks that opened at 0.
                If TickerOpenPrice <> 0 Then
                    Ws.Cells(SummaryTableIndex, "K").Value = FormatPercent((TickerClosePrice - TickerOpenPrice) / TickerOpenPrice, 2)
                Else
                    Ws.Cells(SummaryTableIndex, "K").Value = Null
                End If
                'Deliver the Total Stock Volume for each ticker into the Summary table.
                Ws.Cells(SummaryTableIndex, "L").Value = TotalStockVolume
                'Green for yearly change that goes up.
                If Ws.Cells(SummaryTableIndex, "J").Value > 0 Then
                    Ws.Cells(SummaryTableIndex, "J").Interior.ColorIndex = 4
                Else
                    'Red for yearly change that goes down.
                    Ws.Cells(SummaryTableIndex, "J").Interior.ColorIndex = 3
                End If
                'Summarizing the Summary and Providing Conclusions:
                'Challenge: Locates the largest percentage increase of all the tickers.
                If Ws.Cells(SummaryTableIndex, "K").Value > PercentIncrease Then
                    PercentIncrease = Ws.Cells(SummaryTableIndex, "K").Value
                    PercentIncreaseTotal = Ws.Cells(SummaryTableIndex, "I").Value
                End If
                'Challenge: Locates the largest percentage decrease of all the tickers.
                If Ws.Cells(SummaryTableIndex, "K").Value < PercentDecrease Then
                    PercentDecrease = Ws.Cells(SummaryTableIndex, "K").Value
                    PercentDecreaseTotal = Ws.Cells(SummaryTableIndex, "I").Value
                End If
                'Challenge: Locates the greatest total volume of all the tickers.
                If Ws.Cells(SummaryTableIndex, "L").Value > GreatestTotalVolume Then
                    GreatestTotalVolume = Ws.Cells(SummaryTableIndex, "L").Value
                    GreatestTotalVolumeTicker = Ws.Cells(SummaryTableIndex, "I").Value
                End If
                'Resets the opening conditions for the next ticker.
                TickerOpenPrice = Ws.Cells(RowIndex + 1, "C").Value
                TotalStockVolume = 0
                SummaryTableIndex = SummaryTableIndex + 1
            End If
    Next RowIndex
    'Delivers the data for the challenges and their coresponding ticker symbols.
    Ws.Cells(2, "P").Value = PercentIncreaseTotal
    Ws.Cells(2, "Q").Value = (FormatPercent(PercentIncrease, 2))
    Ws.Cells(3, "P").Value = PercentDecreaseTotal
    Ws.Cells(3, "Q").Value = (FormatPercent(PercentDecrease, 2))
    Ws.Cells(4, "P").Value = GreatestTotalVolumeTicker
    Ws.Cells(4, "Q").Value = GreatestTotalVolume
    'Formats the column widths because I am obsessive about that sort of thing.
    Ws.Columns("A:Q").AutoFit
'Go to the next worksheet and do it all again.
Next Ws
End Sub
