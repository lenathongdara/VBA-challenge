Sub multiple_year_stock_analysis():

'Declare variables
    Dim i As Long
    Dim j As Long
    Dim YearlyChange As Double
    Dim TotalStockVolume As Long
    Dim TickerAmount As Long
    Dim PercentChng As Double
    Dim GreatestInc As Double
    Dim GreatestDec As Double
    Dim GreatestTotVol As Double
    Dim LastRowA As Long
    Dim LastRowI As Long

'Create Titles for cells
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"

'Set initial values
    j = 0
    total = 0
    change = 0
    start = 2

'Formula to find the last row
lastrow = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    total = total + Cells(i, 7).Value

    If total = 0 Then
        Range("I" & 2 + j).Value = Cells(i, 1).Value
        Range("J" & 2 + j).Value = 0
        Range("K" & 2 + j).Value = "%" & 0
        Range("L" & 2 + j).Value = 0

    Else
        If Cells(start, 3) = 0 Then
    For find_value = start To i
    If Cells(find_value, 3).Value <> 0 Then
        start = find_value
    Exit For
    End If
    Next find_value
    End If

change = (Cells(i, 6) - Cells(start, 3))
percentChange = change / Cells(start, 3)

'Begin the next ticker for stock
start = i + 1
                
Range("I" & 2 + j).Value = Cells(i, 1).Value
Range("J" & 2 + j).Value = change
Range("J" & 2 + j).NumberFormat = "0.00"
Range("K" & 2 + j).Value = percentChange
Range("K" & 2 + j).NumberFormat = "0.00%"
Range("L" & 2 + j).Value = total

'Change (+) colors green and (-) colors red
Select Case change
Case Is > 0
    Range("J" & 2 + j).Interior.ColorIndex = 4
Case Is < 0
    Range("J" & 2 + j).Interior.ColorIndex = 3
Case Else
    Range("J" & 2 + j).Interior.ColorIndex = 0
End Select
End If

'Begin the next new ticker
total = 0
change = 0
days = 0
j = j + 1

'Add results if ticker is the same
Else
total = total + Cells(i, 7).Value
End If
Next i

'Caluclate the values
Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100
Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastrow))

'Do not count the header row
final_greatest_increase = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
final_greatest_decrease = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
final_greatest_total_volume = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)

'Final results for the Greatest % of Increase, Greatest % of Decrease, and Greatest Total Volume
Range("P2") = Cells(final_greatest_increase + 1, 9)
Range("P3") = Cells(final_greatest_decrease + 1, 9)
Range("P4") = Cells(final_greatest_total_volume + 1, 9)

End Sub


