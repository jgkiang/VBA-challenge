Sub StockAnalysis()

Dim TotalVolume As Long
Dim Ticker As Long
Dim YearlyChange As Double
Dim PercentChange As Double
Dim CurrentOpen As Double

TotalVolume = 0
Ticker = 2
CurrentOpen = Cells(2, 3)

For i = 2 To 330
    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
       TotalVolume = Cells(i, 7).Value + TotalVolume

    Else
        YearlyChange = Cells(i, 3) - Cells(i, 6)
            If YearlyChange > 0 Then
            Cells(Ticker, 10).Interior.ColorIndex = 4
            Else
            Cells(Ticker, 10).Interior.ColorIndex = 3
            End If
        PercentChange = YearlyChange / CurrentOpen
        TotalVolume = Cells(i, 7).Value + TotalVolume
        Cells(Ticker, 9).Value = Cells(i, 1).Value
        Cells(Ticker, 10).Value = YearlyChange
        Cells(Ticker, 11).Value = PercentChange
        Cells(Ticker, 11).NumberFormat = "0.00%"
        Cells(Ticker, 12).Value = TotalVolume
        TotalVolume = 0
        Ticker = Ticker + 1
        CurrentOpen = Cells(i + 1, 3)
    
    End If

Next i

End Sub
