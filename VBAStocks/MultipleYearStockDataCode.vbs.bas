Attribute VB_Name = "Module1"
Sub StockInfo()

'defined variables for the module
Dim i As Long
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalSVolume As Double
Dim Column As Integer
Dim Ticker As String
Dim Table As Integer
Dim StartRow As Long
Dim LastRow As Long
Dim ws As Worksheet
 
 For Each ws In Worksheets

'values of the variables to be used
Column = 1
Table = 2
i = 2
Ticker = ws.Cells(i, 1).Value
OpeningPrice = ws.Cells(i, 3).Value
StartRow = 2

'create column titles for ticker, yearly change, percent change, and total stock volume

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Create a script that will loop through all the stocks for one year and output the info for: ticker symbol
'yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow
    If ws.Cells(i + 1, Column).Value <> ws.Cells(i, Column).Value Then
        ClosingPrice = ws.Cells(i, 6).Value
        TotalSVolume = ws.Application.Sum(Range(Cells(StartRow, 7), Cells(i, 7)))
        YearlyChange = ClosingPrice - OpeningPrice
        If OpeningPrice <> 0 Then PercentChange = (ClosingPrice - OpeningPrice) / OpeningPrice
       ws.Cells(Table, 9).Value = Ticker
       ws.Cells(Table, 10).Value = YearlyChange
       ws.Cells(Table, 11).Value = PercentChange
       ws.Cells(Table, 12).Value = TotalSVolume
       Table = Table + 1
       TotalSVolume = 0
       Ticker = ws.Cells(i + 1, 1).Value
       StartRow = i + 1
       OpeningPrice = ws.Cells(i + 1, 3).Value
    End If
Next i

'conditional formatting to highlight positive change in green and negative change in red in the YearlyChange column

Z = ws.Cells(Rows.Count, 9).End(xlUp).Row
With ws.Range("j" & 2 & ":j" & Z).FormatConditions.Add(xlCellValue, xlGreater, "=0")
        .Interior.Color = vbGreen
    End With

    With ws.Range("j" & 2 & ":j" & Z).FormatConditions.Add(xlCellValue, xlLess, "=0")
        .Interior.Color = vbRed
    End With
ws.Columns("l:l").NumberFormat = "#,##0_);[Red](#,##0)"
ws.Columns("k").NumberFormat = "0.00%"

'do this for all ws in worksheets

Next ws

End Sub
