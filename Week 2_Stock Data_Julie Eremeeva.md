Sub Stocks1()
'
'Looping through all tabs
Dim ws As Worksheet
For Each ws In Worksheets

'Defining values

Dim Ticker As String
Dim OpenPrice As Double
    
Dim Volume As Long
Dim TickerUnique As String
Dim YearChange As Double
Dim PercentChange As Double
Dim TotalVolume As Long

Dim LastRow As Long

'Creating cell names

 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Total Stock Volume"
 ws.Cells(1, 16).Value = "Ticker"
 ws.Cells(1, 17).Value = "Value"
 ws.Cells(2, 15).Value = "Greatest % Increase"
 ws.Cells(3, 15).Value = "Greatest % Decrease"
 ws.Cells(4, 15).Value = "Greatest Total Volume"


'Defining a row for the TickerUnique column and close price

Dim TickerRow As Long
    TickerRow = "2"

Dim j As Long
    j = 2
    
'Determining the Last Row

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

'Part 1: Outputting the Ticker Sybmol

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    TickerUnique = ws.Cells(i, 1).Value
    ws.Cells(TickerRow, "I").Value = TickerUnique
    
'Part 2: Yearly Change

    YearChange = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
    ws.Cells(TickerRow, "J").Value = YearChange
    
'Part 3: Conditional formatting
   

    If ws.Cells(TickerRow, "J").Value < 0 Then
    
    ws.Cells(TickerRow, "J").Interior.ColorIndex = 3
                
    Else
    
    ws.Cells(TickerRow, "J").Interior.ColorIndex = 4
    
    End If
    
    If ws.Cells(j, 3).Value <> 0 Then
    
        PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
        
        ws.Cells(TickerRow, "K").Value = Format(PercentChange, "Percent")
                    
    Else
                    
        ws.Cells(TickerRow, "K").Value = Format(0, "Percent")
                    
    End If
    
    ws.Cells(TickerRow, "L").Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

    TickerRow = TickerRow + 1
    
    j = i + 1

    End If
    
Next i

'Part 4: Greatest % Inrease, Greatest % Descrease, Greatest Total Volume
'Defining values

Dim PerIncr As Double
PerIncr = ws.Cells(2, "K").Value

Dim PerDecr As Double
PerDecr = ws.Cells(2, "K").Value

Dim TotalVol As Double
TotalVol = ws.Cells(2, "L").Value

'Defining the last row for the ticker unique column to calculate greatest values

Dim LastRowTicker As Long

LastRowTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Three loops to find greatest increase/decrease/volume

For i = 2 To LastRowTicker

    If ws.Cells(i, 11).Value > PerIncr Then
    PerIncr = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).Value = PerIncr
                
    Else
                
    PerIncr = PerIncr

    End If

Next i

For i = 2 To LastRowTicker

    If ws.Cells(i, 11).Value < PerDecr Then
    PerDecr = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 17).Value = PerDecr
                
    Else
                
    PerDecr = PerDecr

    End If

Next i

For i = 2 To LastRowTicker

    If ws.Cells(i, 12).Value > TotalVol Then
    TotalVol = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = TotalVol
                
    Else
                
    TotalVol = TotalVol

    End If

'Formatting the cells as %

ws.Cells(2, 17).Value = Format(PerIncr, "Percent")
ws.Cells(3, 17).Value = Format(PerDecr, "Percent")

Next i

ws.Columns("A:Z").AutoFit

Next ws

'
End Sub