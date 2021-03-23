Sub AnnualStockChange():

'Declaring variables'
Dim ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim volume As Double
Dim summary_table_row As Integer

''''''Creating Summary Table''''''

Range("J1").Value = "Summary Table"
Range("J2").Value = "Ticker"
Range("K2").Value = "Annual Change"
Range("L2").Value = "Percent Change"
Range("M2").Value = "Total Stock Volume"

summary_table_row = 3

'Using this column to store first open price of each stock'
Columns("N").Hidden = True

volume = 0
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Chose to not use these column names to avoid taking up space, but left them as a comment to simplify code readability'
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
        Cells(summary_table_row, 12).Value = (closePrice / openPrice)
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

End Sub