Sub StockAnalysis()

    ' Set variables for holding the ticker name, ticker volume, etc
    Dim ws As Worksheet
    Dim TickerName As String
    Dim TickerVol As Double
    Dim TickerRow As Integer
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim ConditionalRowFormat As Double
    Dim MaxVolume As Double ' To store the greatest total volume
    Dim MaxVolumeTicker As String
    Dim maxPercentChange As Double ' To store the greatest percentage change
    Dim maxPercentChangeTicker As String ' To store the ticker associated with the greatest percentage change
    Dim maxPercentDecrease As Double ' To store the greatest percentage decrease
    Dim maxPercentDecreaseTicker As String ' To store the ticker associated with the greatest percentage decrease
    
    ' Initialize variables
    MaxVolume = 0 ' Initialize MaxVolume
    maxPercentChange = 0 ' Initialize maxPercentChange
    maxPercentDecrease = 0 ' Initialize maxPercentDecrease
    
    For Each ws In ThisWorkbook.Worksheets
        TickerVol = 0
        TickerRow = 2
        
        ' Set up worksheet headers
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(1, 15) = "Ticker"
        ws.Cells(1, 16) = "Value"
        
        ' Set up labels for calculations
        ws.Cells(2, 14) = "Greatest % Increase"
        ws.Cells(3, 14) = "Greatest % Decrease"
        ws.Cells(4, 14) = "Greatest Total Volume"
        
        ' Find the last row and conditional format row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ConditionalRowFormat = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                
                TickerName = ws.Cells(i, 1)
                TickerVol = TickerVol + ws.Cells(i, 7)
                
                ws.Range("I" & TickerRow) = TickerName
                ws.Range("L" & TickerRow) = TickerVol
                
                ClosePrice = ws.Cells(i, 6)
                YearlyChange = ClosePrice - OpenPrice
                ws.Range("J" & TickerRow) = YearlyChange
                
                If (OpenPrice = 0) Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice
                End If
                
                ' Update maxPercentChange and associated ticker
                If PercentChange > maxPercentChange Then
                    maxPercentChange = PercentChange
                    maxPercentChangeTicker = TickerName
                End If
                
                ' Update maxPercentDecrease and associated ticker
                If PercentChange < maxPercentDecrease Then
                    maxPercentDecrease = PercentChange
                    maxPercentDecreaseTicker = TickerName
                End If
                
                ws.Range("K" & TickerRow) = PercentChange
                ws.Range("K" & TickerRow).NumberFormat = "0.00%"
                
                TickerRow = TickerRow + 1
                TickerVol = 0
                OpenPrice = ws.Cells(i + 1, 3)
                
            Else
                TickerVol = TickerVol + ws.Cells(i, 7)
            End If
        Next i
        
        ' Apply conditional formatting
        For j = 2 To ConditionalRowFormat
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
        ' Calculate the greatest total volume
        Dim currentMaxVolume As Double
        currentMaxVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
        If currentMaxVolume > MaxVolume Then
            MaxVolume = currentMaxVolume
            MaxVolumeTicker = TickerName
        End If
        
        ' Store the greatest percentage change and associated ticker
        ws.Range("P2").Value = maxPercentChange
        ws.Range("O2").Value = maxPercentChangeTicker
        
        ' Store the greatest percentage decrease and associated ticker
        ws.Range("P3").Value = maxPercentDecrease
        ws.Range("O3").Value = maxPercentDecreaseTicker
        
        ' Store the greatest total volume and associated ticker
        ws.Range("P4").Value = MaxVolume
        ws.Range("O4").Value = MaxVolumeTicker
    Next ws
End Sub

