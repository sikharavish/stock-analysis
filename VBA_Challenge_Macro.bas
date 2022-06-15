Attribute VB_Name = "VBAChallenge"

Sub VBAChallenge()

Dim startTime As Single
Dim endTime As Single

yearValue = InputBox("What year would you like to run the analysis on?", "Stock Analysis", "Any Year Stock Analysis")

startTime = Timer

Worksheets("VBAChallenge").Activate

Range("A1").Value = "All Stocks (" + yearValue + ")"

'Create a header row

Cells(3, 1).Value = "Tickers"
Cells(3, 2).Value = "Ticker Volumes"
Cells(3, 3).Value = "TickerStartingPrices"
Cells(3, 4).Value = "TickerEndingPrices"
Cells(3, 5).Value = "Return"

Dim tickers(11) As String

tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"

'Initialize variables for starting price and ending price

Dim TickerStartingPrices As Single
Dim TickerEndingPrices As Single

Worksheets(yearValue).Activate

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    
    Worksheets(yearValue).Activate
    
    For j = 2 To RowCount
    

        If Cells(j, 1).Value = ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value
    End If

        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        TickerStartingPrices = Cells(j, 6).Value
    End If

        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        TickerEndingPrices = Cells(j, 6).Value
    End If
Next j
    Worksheets("VBAChallenge").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = TickerStartingPrices
    Cells(4 + i, 4).Value = TickerEndingPrices
    Cells(4 + i, 5).Value = TickerEndingPrices / TickerStartingPrices - 1
    
Next i


 'Formatting
    Worksheets("VBAChallenge").Activate
    Range("A3:E3").Font.Bold = True
    Range("A3:E3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:D15").NumberFormat = "#,##0"
    Range("E4:E15").NumberFormat = "0.0%"
    Columns("A:E").AutoFit
    
dataRowStart = 4
dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 5) > 0 Then

            'Color the cell green
            Cells(i, 5).Interior.Color = vbGreen

        ElseIf Cells(i, 5) < 0 Then

            'Color the cell red
            Cells(i, 5).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 5).Interior.Color = xlNone

        End If

    Next i
    
    endTime = Timer
    MsgBox "This code ran in" & (endTime - startTime) & "seconds for the year" & (yearValue)

End Sub
