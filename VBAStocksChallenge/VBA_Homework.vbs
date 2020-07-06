Sub VBA_Challenge():

Dim ws as Worksheet

For each ws in Worksheets

ws.Activate

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"

    Dim Ticker_Name as String
    Dim YearClosing as double
    Dim YearOpening as double
    YearOpening = Cells(2, 3).Value
    Dim YearlyChange as double
    Dim PercentChange as double
    Dim StockVolume as double
    StockVolume = 0

    Dim Summary_Table_Row as Integer
    Summary_Table_Row = 2

      'Finds the last non-blank cell in Column
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'Establish For Loop
    For i = 2 to LastRow

        If Cells(i, 1).Value <> Cells(i + 1, 1).Value then
            
            Ticker_Name = Cells(i, 1).Value
            YearClosing = Cells(i, 6).Value

            YearlyChange = YearClosing - YearOpening

            If YearOpening = 0 and YearClosing = 0 then
                PercentChange = 0
            'PercentChange = (YearClosing - YearOpening) / YearClosing

            ElseIf YearClosing >= YearOpening then
                PercentChange = (YearClosing / YearOpening) - 1
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            'PercentChange = (YearClosing - YearOpening) / YearClosing

            Else
                PercentChange = (YearClosing / YearOpening) - 1
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

            'PercentChange = (YearClosing / YearOpening
            
            End If
        
        'insert Ticker_Name into Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker_Name

        'insert YearlyChange into Summary Table
        Range("J" & Summary_Table_Row).Value = YearlyChange

        'insert PercentChange into Summary Table
        Range("K" & Summary_Table_Row).Value = PercentChange
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

        'insert StockVolume into Summary Table
        Range("L" & Summary_Table_Row).Value = StockVolume
        
        Summary_Table_Row = Summary_Table_Row + 1

        StockVolume = 0
        YearlyChange = 0
        PercentChange = 0
        YearOpening = Cells(i + 1, 3).Value

        Else
            
            StockVolume = StockVolume + Cells(i, 7).Value

        End If
    
    Next i

    Dim LastRowSummary As Long
    LastRowSummary = Cells(Rows.Count, 9).End(xlUp).Row
    Dim PercentIncrease as double
    Dim PercentDecrease as double
    Dim TickerMin as String
    Dim TickerMax as String
    Dim TotalVolume as double
    Dim TickerVol as String
    PercentDecrease = Cells(2, 11).Value
    PercentIncrease = Cells(2, 11).Value
    TotalVolume = Cells(2, 12).Value
    



    For i = 3 to LastRowSummary

            If Cells(i, 11).Value < PercentDecrease then
                PercentDecrease = Cells(i, 11).Value
                TickerMin = Cells(i, 9).Value
                'Range(Q3).Value = PercentDecrease
                'Range(P3).Value = TickerMin
                'Range(Q3).NumberFormat = "0.00%"

            End If

            If Cells(i, 11).Value > PercentIncrease then
                PercentIncrease = Cells(i, 11).Value
                TickerMax = Cells(i, 9).Value
                'Range(Q2).Value = PercentIncrease
                'Range(P2).Value = TickerMax
                'Range(Q2).NumberFormat = "0.00%"

            End If 

            If Cells(i, 12).Value > TotalVolume then
                TotalVolume = Cells(i, 12).Value
                TickerVol = Cells(i, 9).Value
                'Range(Q4).Value = TotalVolume
                'Range(P4).Value = TickerVol

            End If
        

    Next i

    Cells(3, 17).Value = PercentDecrease
    Cells(3, 16).Value = TickerMin
    Cells(3, 17).NumberFormat = "0.00%"

    Cells(2, 17).Value = PercentIncrease
    Cells(2, 16).Value = TickerMax
    Cells(2, 17).NumberFormat = "0.00%"

    Cells(4, 17).Value = TotalVolume
    Cells(4, 16).Value = TickerVol

Next ws

End Sub
