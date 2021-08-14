Sub MultipleYearStock()

Dim ws As Worksheet
For Each ws In Worksheets

Dim lastRow, newNextRow, I As Long
Dim counter As Integer
Dim firstOpen, lastClose As Double
Dim ticker As String
Dim yearlyChange, percentChange, totalVol As Double


'count the number of rows
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'when add ticker to a new row, use this newNextRow to avoid blank cells
newNextRow = 2
'a variable to store total vol for a certain ticker
totalVol = 0
' define a counter to record how many days in that year
counter = 0
For I = 2 To lastRow

'sum up vol in for loop
totalVol = totalVol + ws.Cells(I, 7).Value
'calculate days in that year
counter = counter + 1

'to check if next row still the same ticker or not
If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
'to store the same type of ticker in this variable
ticker = ws.Cells(I, 1).Value
firstOpen = ws.Cells(I - counter + 1, 3).Value
lastClose = ws.Cells(I, 6).Value
yearlyChange = lastClose - firstOpen

' sometimes the price of first day of the year is 0, and this may cause error
If firstOpen = 0 Then
percentChange = 0
ElseIf firstOpen <> 0 Then
percentChange = yearlyChange / firstOpen
End If

'assign these data into new columns
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

'assign summary information of ticker in new columns
ws.Cells(newNextRow, 9) = ticker
ws.Cells(newNextRow, 10) = yearlyChange

'conditional formatting the positive and negative yearly change
If (yearlyChange > 0) Then
ws.Cells(newNextRow, 10).Interior.ColorIndex = 4
ElseIf (yearlyChange < 0) Then
ws.Cells(newNextRow, 10).Interior.ColorIndex = 3
ElseIf (yearlyChange = 0) Then
ws.Cells(newNextRow, 10).Interior.ColorIndex = 2
End If

ws.Cells(newNextRow, 11) = percentChange
ws.Cells(newNextRow, 12) = totalVol

'move to next row to store another type of ticker
newNextRow = newNextRow + 1
'reset the total vol for next type of ticker
totalVol = 0
counter = 0
End If
Next I


'format the column of percent change in x.xx%
ws.Columns("K").NumberFormat = "0.00%"
ws.Columns("A:Z").AutoFit

Next ws


'Challenge part

Dim maxChange, minChange, maxVol As Double
Dim comparemaxC, compareminC, comparemaxV As Double
'define three variables to hold ticker name
Dim maxCT, minCT, maxVT As String

'initial assign numbers into compare variable to allow multi-page comparison in the following steps
comparemaxC = 0
compareminC = 1000000000
comparemaxV = 0


Dim ws2 As Worksheet
For Each ws2 In Worksheets


'pick up max and min change in the summarised info in each sheet
'add titles
ws2.Range("P1") = "Ticker"
ws2.Range("Q1") = "Value"
ws2.Range("O2") = "Greatest % Increase"
ws2.Range("O3") = "Greatest % Decrease"
ws2.Range("O4") = "Greatest Total Volume"

maxChange = Application.WorksheetFunction.Max(ws2.Range("K2:K" & newNextRow))
'compare is the current max % change bigger than older stored biggest % change
'if bigger than older one, then substitute the new biggest into comparemaxC to allow comparison in next worksheet
If maxChange > comparemaxC Then
comparemaxC = maxChange
maxCT = ws2.Range("I" & (WorksheetFunction.Match(maxChange, ws2.Range("K2:K" & newNextRow), 0) + 1)).Value
End If
ws2.Range("Q2") = maxChange
ws2.Range("Q2").NumberFormat = "0.00%"
ws2.Range("P2") = ws2.Range("I" & (WorksheetFunction.Match(maxChange, ws2.Range("K2:K" & newNextRow), 0) + 1)).Value

minChange = Application.WorksheetFunction.Min(ws2.Range("K2:K" & newNextRow))
'compare is the current min % change smaller than older stored smallest % change
'if smaller than older one, then substitute the new smallest into compareminC to allow comparison in next worksheet
If minChange < compareminC Then
compareminC = minChange
minCT = ws2.Range("I" & (WorksheetFunction.Match(minChange, ws2.Range("K2:K" & newNextRow), 0) + 1)).Value
End If
ws2.Range("Q3") = minChange
ws2.Range("Q3").NumberFormat = "0.00%"
ws2.Range("P3") = ws2.Range("I" & (WorksheetFunction.Match(minChange, ws2.Range("K2:K" & newNextRow), 0) + 1)).Value

'same idea applied here as comparemaxC
maxVol = Application.WorksheetFunction.Max(ws2.Range("L2:L" & newNextRow))
If maxVol > comparemaxV Then
comparemaxV = maxVol
maxVT = ws2.Range("I" & (WorksheetFunction.Match(maxVol, ws2.Range("L2:L" & newNextRow), 0) + 1)).Value
End If
ws2.Range("Q4") = maxVol
ws2.Range("P4") = ws2.Range("I" & (WorksheetFunction.Match(maxVol, ws2.Range("L2:L" & newNextRow), 0) + 1)).Value

ws2.Columns("A:Z").AutoFit

Next ws2

'Assign the compared results for max/min % change and max Vol ticker in the current opening worksheet
'Range("P6") = "Ticker"
'Range("Q6") = "Value"
'Range("O7") = "Greatest % Increase through multi-sheets"
'Range("O8") = "Greatest % Decrease through multi-sheets"
'Range("O9") = "Greatest Total Volume through multi-sheets"
'Range("P7") = maxCT
'Range("P8") = minCT
'Range("P9") = maxVT
'Range("Q7") = comparemaxC
'Range("Q7").NumberFormat = "0.00%"
'Range("Q8") = compareminC
'Range("Q8").NumberFormat = "0.00%"
'Range("Q9") = comparemaxV
'
'Columns("A:Z").AutoFit

End Sub
