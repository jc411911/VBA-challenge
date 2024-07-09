Sub QaurterlyStock()
    Dim ws As Worksheet
    Dim summary_ticker_row As Integer
    Dim open_price As Double, close_price As Double
    Dim quarterly_change As Double, percent_change As Double
    Dim lastrow As Long
    Dim tickervolume As Double
    Dim greatest_increase As Double, greatest_decrease As Double, greatest_volume As Double
    Dim ticker_greatest_increase As String, ticker_greatest_decrease As String, ticker_greatest_volume As String
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Q*" Then
            ws.Activate
            tickervolume = 0
            summary_ticker_row = 2
            open_price = Cells(2, 3).Value
            greatest_increase = -99999
            greatest_decrease = 99999
            greatest_volume = 0
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Quarterly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To lastrow
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    tickername = Cells(i, 1).Value
                    tickervolume = tickervolume + Cells(i, 7).Value

                    Range("I" & summary_ticker_row).Value = tickername
                    Range("L" & summary_ticker_row).Value = tickervolume
                    close_price = Cells(i, 6).Value

                    quarterly_change = close_price - open_price
                    Range("J" & summary_ticker_row).Value = quarterly_change

                    If open_price = 0 Then
                        percent_change = 0
                    Else
                        percent_change = quarterly_change / open_price
                    End If

                    Range("K" & summary_ticker_row).Value = percent_change
                    Range("K" & summary_ticker_row).NumberFormat = "0.00%"

                    If percent_change > greatest_increase Then
                        greatest_increase = percent_change
                        ticker_greatest_increase = tickername
                    End If

                    If percent_change < greatest_decrease Then
                        greatest_decrease = percent_change
                        ticker_greatest_decrease = tickername
                    End If

                    If tickervolume > greatest_volume Then
                        greatest_volume = tickervolume
                        ticker_greatest_volume = tickername
                    End If

                    summary_ticker_row = summary_ticker_row + 1
                    tickervolume = 0
                    open_price = Cells(i + 1, 3)
                Else
                    tickervolume = tickervolume + Cells(i, 7).Value
                End If
            Next i
            lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
            For i = 2 To lastrow_summary_table
                If Cells(i, 10).Value = 0 Then
                    Cells(i, 10).Interior.ColorIndex = -4142
                ElseIf Cells(i, 10).Value > 0 Then
                    Cells(i, 10).Interior.ColorIndex = 10
                Else
                    Cells(i, 10).Interior.ColorIndex = 3
                End If
            Next i
            Range("Q1").Value = "Ticker"
            Range("R1").Value = "Value"
            Range("P2").Value = "Greatest % Increase"
            Range("Q2").Value = ticker_greatest_increase
            Range("R2").Value = greatest_increase
            Range("R2").NumberFormat = "0.00%"
            Range("P3").Value = "Greatest % Decrease"
            Range("Q3").Value = ticker_greatest_decrease
            Range("R3").Value = greatest_decrease
            Range("R3").NumberFormat = "0.00%"
            Range("P4").Value = "Greatest Total Volume"
            Range("Q4").Value = ticker_greatest_volume
            Range("R4").Value = greatest_volume
        End If
    Next ws
End Sub


