Sub stockticker()

'create variables'
'1create variable for ticker
Dim ticker As String

'2create a variable for the row that the stock information will be reported to on the ticker
Dim Tickerrow As Integer

'3create variable for Quarterly change
Dim Quarterlychange As Double

'4create variable for percent change
Dim PercentChange As Double

'5create variable for the total volume of stocks traded
Dim TotalVolume As Double

'6create variable to keep track of the rowcount
Dim RowCount As Long

'7create Variables for calculating quarterly and percent change
Dim Startprice As Double
Dim Endprice As Double

'8create Variables for counting lowest percent change, highest percent change, and max volume
Dim maxvalue As Double
Dim lowestpercent As Double
Dim highestpercent As Double

'9Create variable for Rowcount for finding max and min
Dim finalrowcount As Long

'10Create variable for formatting each worksheet
For Each ws In Worksheets
TotalVolume = 0
Tickerrow = 2
PercentChange = 0
Quarterlychange = 0
Startprice = ws.Range("C2").Value

'Range values for column/row Headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quarterly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Finding rowcount for each sheet
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To RowCount
'If statement to check to see if the ticker for the new row is the same as the preceding row
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'If the row is different, add a new line to the ticker row to indicate new stock
ticker = ws.Cells(i, 1).Value

'Finding Endprice at end of year
Endprice = ws.Cells(i, 6).Value

'Finding Quarterlychange
Quarterlychange = Endprice - Startprice

'Finding percentchange
PercentChange = (Quarterlychange / Startprice)

'Begin counting total volume
TotalVolume = TotalVolume + ws.Cells(i, 7).Value

'Print results in the excel worksheet
ws.Range("I" & Tickerrow).Value = ticker
ws.Range("L" & Tickerrow).Value = TotalVolume
ws.Range("J" & Tickerrow).Value = Quarterlychange
ws.Range("K" & Tickerrow).Value = PercentChange

'set percentage formatting for percentchange
ws.Range("K" & Tickerrow).NumberFormat = "0.00%"

'Setting conditional formatting
If Quarterlychange < 0 Then
ws.Range("J" & Tickerrow).Interior.ColorIndex = 3
'ws.Range("K" & Tickerrow).Interior.ColorIndex = 3
Else
ws.Range("J" & Tickerrow).Interior.ColorIndex = 4
'ws.Range("K" & Tickerrow).Interior.ColorIndex = 4
End If

'Settinh startprice to new value
Startprice = ws.Cells(i + 1, 3).Value

'setting endprice back to 0
Endprice = 0

'setting Quarterlychange back to 0
Quarterlychange = 0

'setting percentchange back to 0
PercentChange = 0

'Add 1 to the ticker row so that next stock doesnt override data
Tickerrow = Tickerrow + 1

'Reset totalvolume count so that the total for the next stock starts at 0
TotalVolume = 0

Else
'If the If statement determines that the next row is same stock, just update the total
TotalVolume = TotalVolume + ws.Cells(i, 7).Value

End If
'move to next i
Next i

'Find the rowcount for each unique stock on each page
finalrowcount = Cells(Rows.Count, "I").End(xlUp).Row

'initialize variables
maxvalue = 0
lowestpercent = 0
highestpercent = 0

'For loop for calculating min/max
For j = 2 To finalrowcount
'Calculation for determining max value of volume.
'Comparing each row to the previous max, starting at 0.
'If the value is greater than previous max, then it is printed on the sheet and becomes the new maximum value
'If another value comes along that is higher, the previous printout is overridden.
If ws.Cells(j, 12).Value > maxvalue Then
ws.Range("P4") = ws.Cells(j, 9).Value
ws.Range("Q4") = ws.Cells(j, 12).Value

'Keep track of new max value
maxvalue = ws.Cells(j, 12).Value
Else

'We dont care if the value isnt the max, so nothing here
End If

'Calculation for finding lowest percentage
If ws.Cells(j, 11).Value < lowestpercent Then
ws.Range("P3") = ws.Cells(j, 9).Value
ws.Range("Q3") = ws.Cells(j, 11).Value

'Converts format to percentage
ws.Range("Q3").NumberFormat = "0.00%"
lowestpercent = ws.Cells(j, 11).Value
Else
End If

'Calculation for finding highest percentage
If ws.Cells(j, 11).Value > highestpercent Then
ws.Range("P2") = ws.Cells(j, 9).Value
ws.Range("Q2") = ws.Cells(j, 11).Value
ws.Range("Q2").NumberFormat = "0.00%"
highestpercent = ws.Cells(j, 11).Value
Else
End If

Next j
Next ws
End Sub



