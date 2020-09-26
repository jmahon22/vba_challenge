VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_data()

Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As Double
Dim TableRow As Integer
Dim OpenPrice As Double
Dim YearlyOpen As Double
Dim YearlyClose As Double
Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestTotalVolume As Double

'setting up headers/table

For Each ws In ThisWorkbook.Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
Next

For Each ws In ThisWorkbook.Worksheets

TotalVolume = 0
TableRow = 2
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
YearlyOpen = ws.Cells(2, 3).Value

For i = 2 To LastRow
    'checking to see if in the same ticker, if not:
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'ticker symbol and total volume
        Ticker = ws.Cells(i, 1).Value
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            ws.Range("I" & TableRow).Value = Ticker
            ws.Range("L" & TableRow).Value = TotalVolume
        'yearly change
        YearlyClose = ws.Cells(i, 6).Value
        YearlyChange = YearlyClose - YearlyOpen
            ws.Range("J" & TableRow).Value = YearlyChange
            
            'conditional formatting
            If ws.Range("J" & TableRow).Value > 0 Then
                ws.Range("J" & TableRow).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & TableRow).Value < 0 Then
                ws.Range("J" & TableRow).Interior.ColorIndex = 3
            End If
    'percent change
    If (YearlyOpen = 0 And YearlyClose = 0) Then
        PercentChange = 0

    ElseIf (YearlyOpen = 0 And YearlyClose <> 0) Then
        PercentChange = -1

    Else: PercentChange = (YearlyChange / YearlyOpen)
        ws.Range("K" & TableRow).Value = PercentChange
        ws.Range("K" & TableRow).NumberFormat = "0.00%"

    End If
    TableRow = TableRow + 1
    TotalVolume = 0
    YearlyOpen = ws.Cells(i + 1, 3).Value

Else
TotalVolume = TotalVolume + ws.Cells(i, 7).Value


End If

Next i

GreatestIncrease = 0
GreatestDecrease = 0
GreatestTotalVolume = 0

LastRows = ws.Cells(Rows.Count, 11).End(xlUp).Row

For j = 2 To LastRows

'greatest increase
    If ws.Range("K" & j).Value > GreatestIncrease Then
    GreatestIncrease = ws.Range("K" & j).Value
    ws.Range("Q2").Value = GreatestIncrease
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("P2").Value = ws.Range("I" & j).Value
    
'greatest decrease
    ElseIf ws.Range("K" & j).Value < GreatestDecrease Then
    GreatestDecrease = ws.Range("K" & j).Value
    ws.Range("Q3").Value = GreatestDecrease
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P3").Value = ws.Range("I" & j).Value
    
    End If
'greatest total volume
    If ws.Range("L" & j).Value > GreatestTotalVolume Then
    GreatestTotalVolume = ws.Range("L" & j).Value
    ws.Range("Q4").Value = GreatestTotalVolume
    ws.Range("P4").Value = ws.Range("I" & j).Value
    
    End If
    

Next j

Next ws

End Sub
