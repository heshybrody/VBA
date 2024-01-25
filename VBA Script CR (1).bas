Attribute VB_Name = "Module1"
Sub alpha_test()

Dim i As Long
Dim lastrow As Long
Dim row As Long
Dim ticker As String
Dim total As Double
Dim first As Double
Dim last As Double

For Each ws In Worksheets

' Assign variables
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
row = 2
total = 0
first = ws.Cells(2, 3)

' Title Labels
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"
ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(4, 15) = "Greatest Total Volume"
ws.Cells(1, 16) = "Ticker"
ws.Cells(1, 17) = "Value"

For i = 2 To lastrow
    ' Add each value in the "vol" column for a certain ticker symbol to find the total volume
    total = total + ws.Cells(i, 7).Value
    
    ' Assign a variable for the ticker symbol
    ticker = ws.Cells(i, 1)
    
    ' For each unique ticker, that is not equal to the one before it, we fill in certain values
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ws.Cells(row, 9) = ticker
    ' 'i' should now represent the last row of a certain ticker symbol, so subtract the first row for that stock to find yearly change
    ws.Cells(row, 10) = ws.Cells(i, 6) - first
    ' Divide yearly change by first open value to find percent change
    ws.Cells(row, 11) = FormatPercent(ws.Cells(row, 10) / first)
    ' Print the total volume in column 12
    ws.Cells(row, 12) = total

    ' Reset variables, 'first' becomes the first open price of the next ticker symbol,
    ' total goes back to zero, and
    ' row adds one so the new printed values don't overlap
    first = ws.Cells(i + 1, 3)
    total = 0
    row = row + 1
    End If
    
    ' Conditional formatting the "Yearly Change" column
    If ws.Cells(i, 10) < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
    ElseIf ws.Cells(i, 10) > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 50
    End If
    
Next i

' Greatest Increase, Decrease, Total Volume
Dim j As Double
Dim k As Integer
Dim lastrow2 As Long

'Set new "lastrow" to find the last row of the 4 new columns
lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).row

For j = 2 To lastrow2
    
    If ws.Cells(j, 11) = WorksheetFunction.Max(ws.Range(ws.Cells(2, 11), ws.Cells(lastrow2, 11))) Then
    ws.Cells(2, 16) = ws.Cells(j, 9)
    ws.Cells(2, 17) = FormatPercent(ws.Cells(j, 11))
    ElseIf ws.Cells(j, 11) = WorksheetFunction.Min(ws.Range(ws.Cells(2, 11), ws.Cells(lastrow2, 11))) Then
    ws.Cells(3, 16) = ws.Cells(j, 9)
    ws.Cells(3, 17) = FormatPercent(ws.Cells(j, 11))
    ElseIf ws.Cells(j, 12) = WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(lastrow2, 12))) Then
    ws.Cells(4, 16) = ws.Cells(j, 9)
    ws.Cells(4, 17) = ws.Cells(j, 12)
    
    End If
    
Next j

Next ws

End Sub
