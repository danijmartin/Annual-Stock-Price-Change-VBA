Sub AnnualStockChange():

'Declaring variables'
Dim ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim volume As Double
Dim summary_table_row As Integer
Dim ws As Worksheet


''''''Adding ability to run on all spreadsheets at once''''''
For Each ws In Worksheets
ws.select
    

''''''Creating Summary Table''''''

Range("J1:M1").Merge
Range("J1:M2").Font.Bold = True
Range("J1:M2").HorizontalAlignment = xlCenter
Range("J1").Value = "Summary Table"
Range("J2").Value = "Ticker"
Range("K2").Value = "Annual Change"
Range("L2").Value = "Percent Change"
Range("M2").Value = "Total Stock Volume"

summary_table_row = 3
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
        Cells(summary_table_row, 10).Value = ticker
        Cells(summary_table_row, 11).Value = (closePrice - openPrice)

        If openPrice = 0 and closePrice = 0 Then
            Cells(summary_table_row, 12).Value = "Price stayed at 0"
        Elseif openPrice = 0 and closeprice <> 0 Then
            Cells(summary_table_row, 12).Value = "Stock opened at zero this year"
        Else
            Cells(summary_table_row, 12).Value = ((closePrice - openPrice) / openPrice)
        End if
        
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

lastrow2 = Cells(Rows.Count, 11).End(xlUp).Row

Range("L:L").NumberFormat = "###.##%"
Range("M:M").NumberFormat = "###,###,###"

For i = 3 To lastrow2

    If Cells(i, 11).Value > 0 Then
        Cells(i, 11).Interior.ColorIndex = 4
    ElseIf Cells(i, 11).Value < 0 Then
        Cells(i, 11).Interior.ColorIndex = 3
    End If

Next i

''''''Adding Greatest Percentage Increase, Percentage Decrease, and Total Volume''''''

Range("Q1").Value = "Ticker"
Range("R1").Value = "Value"
Range("P2").Value = "Greatest Percentage Increase"
Range("P3").Value = "Greatest Percentage Decrease"
Range("P4").Value = "Greatest Total Volume"

Dim percentage_range As Range
Dim volume_range As Range
Dim ticker_range As Range

Set percentage_range = Range("L:L")
Set volume_range = Range("M:M")
Set ticker_range = Range("J:J")

Range("R2").Value = Application.WorksheetFunction.Max(percentage_range)
Range("R3").Value = Application.WorksheetFunction.Min(percentage_range)
Range("R4").Value = Application.WorksheetFunction.Max(volume_range)
Range("Q2").Value = Application.WorksheetFunction.XLookup(Range("R2").Value, percentage_range, ticker_range, 0)
Range("Q3").Value = Application.WorksheetFunction.XLookup(Range("R3").Value, percentage_range, ticker_range, 0)
Range("Q4").Value = Application.WorksheetFunction.XLookup(Range("R4").Value, volume_range, ticker_range, 0)

Range("R2:R3").NumberFormat = "###.##%"
Range("R4").NumberFormat = "###,###,###"

Range("Q1:R1").Font.Bold = True
Range("Q1:R1").HorizontalAlignment = xlCenter
Range("P1:P4").Font.Bold = True
Range("P1:P4").HorizontalAlignment = xlLeft


ws.Cells.EntireColumn.AutoFit
'Using this column to store first open price of each stock'
ws.Columns("N").Hidden = True
Next

End Sub
