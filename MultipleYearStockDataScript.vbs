Sub MultipleYearStockData()
    Dim ws As Worksheet
    Dim TickerSymbol As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim Table As Integer
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotal As Double
    Dim LastRow As Long
    Dim i As Long

    For Each ws In Worksheets
        ' Headers
        With ws
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"
        End With

        ' Initialize Variables
        YearlyChange = 0
        PercentChange = 0
        TotalStockVolume = 0
        Table = 2
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotal = 0

        ' Find the last row of data
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' OpenPrice for the first stock
        OpenPrice = ws.Cells(2, 3).Value

        ' Loop through all rows
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TickerSymbol = ws.Cells(i, 1).Value
                ClosePrice = ws.Cells(i, 6).Value

                ' Yearly Change
                YearlyChange = ClosePrice - OpenPrice
                ws.Range("I" & Table).Value = TickerSymbol
                ws.Range("J" & Table).Value = YearlyChange

                ' Percent Change
                If OpenPrice <> 0 Then
                    PercentChange = (YearlyChange / OpenPrice)
                Else
                    PercentChange = 0
                End If
                ws.Range("K" & Table).Value = PercentChange
                ws.Range("K" & Table).NumberFormat = "0.00%"

                ' Total Stock Volume
                ws.Range("L" & Table).Value = TotalStockVolume

                ' Prepare for next stock
                Table = Table + 1
                TotalStockVolume = 0
                OpenPrice = ws.Cells(i + 1, 3).Value
            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i

        ' Coloring Yearly Change
        For i = 2 To LastRow
            If ws.Range("J" & i).Value < 0 Then
                ws.Range("J" & i).Interior.Color = vbRed
            Else
                ws.Range("J" & i).Interior.Color = vbGreen
            End If
        Next i
        
         'Calculates Table And formatting
         For i = 2 To LastRow
            
            If ws.Range("K" & i).Value > GreatestIncrease Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                GreatestIncrease = ws.Range("Q2").Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
                ws.Range("Q2").NumberFormat = "0.00%"
                
            End If

            If ws.Range("K" & i).Value < GreatestDecrease Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                GreatestDecrease = ws.Range("Q3").Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
                ws.Range("Q3").NumberFormat = "0.00%"
                    
            End If

            If ws.Range("L" & i).Value > GreatestTotal Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                GreatestTotal = ws.Range("Q4").Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
                
            End If
            
        Next i
        
        ' AutoFit Columns
        ws.Columns("I:Q").AutoFit
    
    Next ws
    
End Sub
