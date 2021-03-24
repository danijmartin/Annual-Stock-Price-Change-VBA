Sub AnnualStockChange():

'Declaring variables'
Dim ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim volume As Double
Dim summary_table_row As Integer

''''''Creating Summary Table''''''

Range("J1:M1").Merge
Range("J1:M2").Font.Bold = True
Range("J1:M2").HorizontalAlignment = xlcenter
Range("J1").Value = "Summary Table"
Range("J2").Value = "Ticker"
Range("K2").Value = "Annual Change"
Range("L2").Value = "Percent Change"
Range("M2").Value = "Total Stock Volume"

activesheet.Cells.EntireColumn.AutoFit


summary_table_row = 3

'Using this column to store first open price of each stock'
Columns("N").Hidden = True

volume = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Chose not to use these column names to avoid taking up space, but left them as a comment to simplify code readability'
'tickercol = 1
'opencol = 3
'closedcol = 6
'volumecol = 7
'summarytickercol = 10
'summarychangecol = 11
'summarypercentcol = 12
'summaryvolumecol = 13

For i = 2 To lastrow

ticker = Cells(i, 1).Value
next_ticker = Cells(i + 1, 1).Value

    If ticker <> next_ticker Then
        volume = volume + Cells(i, 7).Value
        ticker = Cells(i, 1).Value
        closePrice = Cells(i, 6).Value
        openPrice = Cells(summary_table_row, 14).Value
        Cells(summary_table_row, 11).Value = (closePrice - openPrice)
        Cells(summary_table_row, 12).Value = ((closePrice / openPrice)-1)
        Cells(summary_table_row, 10).Value = ticker
        Cells(summary_table_row, 13).Value = volume
        volume = 0
        ticker = ""
        summary_table_row = summary_table_row + 1
    Else
        volume = volume + Cells(i, 7).Value
        If Cells(summary_table_row, 14).Value = "" Then
            openPrice = Cells(i, 3).Value
            Cells(summary_table_row, 14).Value = openPrice
        End If
    End If

Next i

''''''Adding Formatting to Summary Table''''''

lastrow2 = cells(rows.count,11).end(xlup).row

Range("L:L").NumberFormat = "###.##%"
Range("M:M").NumberFormat = "###,###,###"

For i = 3 to lastrow2

    If cells(i,11).value > 0 Then
        Cells(i,11).interior.colorindex = 4
    Elseif Cells(i,11).value < 0 Then
        Cells(i,11).interior.colorindex = 3
    End If

Next i

''''''Adding Greatest Percentage Increase, Percentage Decrease, and Total Volume''''''

Range("Q1").value = "Ticker"
Range("R1").value = "Value"
Range("P2").value = "Greatest Percentage Increase"
Range("P3").value = "Greatest Percentage Decrease"
Range("P4").value = "Greatest Total Volume"

Dim percentage_range as range
Dim volume_range as range
Dim ticker_range as range

set percentage_range = Range("L:L")
set volume_range = Range("M:M")
set ticker_range = Range("J:J")

Range("R2").value = application.worksheetfunction.max(percentage_range)
Range("R3").value = application.worksheetfunction.min(percentage_range)
Range("R4").value = application.worksheetfunction.max(volume_range)
Range("Q2").value = application.worksheetfunction.xlookup(Range("R2").value, percentage_range, ticker_range,0)
Range("Q3").value = application.worksheetfunction.xlookup(Range("R3").value, percentage_range, ticker_range,0)
Range("Q4").value = application.worksheetfunction.xlookup(Range("R4").value, volume_range, ticker_range,0)

Range("R2:R3").NumberFormat = "###.##%"
Range("R4").NumberFormat = "###,###,###"

Range("Q1:R1").Font.Bold = True
Range("Q1:R1").HorizontalAlignment = xlcenter
Range("P1:P4").Font.Bold = True
Range("P1:P4").HorizontalAlignment = xlleft

End Sub