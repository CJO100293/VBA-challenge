Sub VBA_challenge()

'Step 1 - Script that loops through all stocks for single year and outputs ticker symbol
'Setting Variables
Dim j As Double
Dim start As Double
Dim change As Double
Dim percentage As Double
Dim rowcount As Long
Dim totalvolume As Double
Dim GreatestIncName As String
Dim GreatestIncVal As Double
Dim GreatestDecName As String
Dim GreatestDecVal As Double
Dim GreatestVolName As String
Dim GreatestVolVal As Double
Dim cs As String
For Each ws In Worksheets
cs = ws.Name

'setting starting values of variables
change = 0
percentage = 0
j = 0
start = 2
rowcount = Worksheets(cs).UsedRange.Rows.Count

'Inserting New Header Columns
Worksheets(cs).Range("I1").Value = "Ticker"
Worksheets(cs).Range("J1").Value = "Yearly Change"
Worksheets(cs).Range("K1").Value = "Percentage Change"
Worksheets(cs).Range("L1").Value = "Total Stock Volume"
Worksheets(cs).Range("O2").Value = "Greatest % Increase"
Worksheets(cs).Range("O3").Value = "Greatest % Decrease"
Worksheets(cs).Range("O4").Value = "Greatest Total Volume"
Worksheets(cs).Range("P1").Value = "Ticker"
Worksheets(cs).Range("Q1").Value = "Value"

'Setting width of the new header columns
Worksheets(cs).Range("I:I").ColumnWidth = 6.86
Worksheets(cs).Range("J:J").ColumnWidth = 12.86
Worksheets(cs).Range("K:K").ColumnWidth = 17.57
Worksheets(cs).Range("L:L").ColumnWidth = 17.57
Worksheets(cs).Range("O:O").ColumnWidth = 20.43
Worksheets(cs).Range("P:P").ColumnWidth = 7
Worksheets(cs).Range("Q:Q").ColumnWidth = 10.71

'Making new header columns have bold font
Worksheets(cs).Range("I1").Font.Bold = True
Worksheets(cs).Range("J1").Font.Bold = True
Worksheets(cs).Range("K1").Font.Bold = True
Worksheets(cs).Range("L1").Font.Bold = True
Worksheets(cs).Range("O2").Font.Bold = True
Worksheets(cs).Range("O3").Font.Bold = True
Worksheets(cs).Range("O4").Font.Bold = True
Worksheets(cs).Range("P1").Font.Bold = True
Worksheets(cs).Range("Q1").Font.Bold = True

'Loop through rows in the column
'Every single time it loops through the rows
'totalvolume = totalvolume + G(r)
'so every time it loops it adds G2, G3, G4, G5 values up
'we reset this once you find a new stock name so it does not get mixed up
'totalvolume = 0  down at the bottom of this main For R loop (rows loop)
        For r = 2 To rowcount
        totalvolume = totalvolume + Worksheets(cs).Cells(r, 7).Value 'PART OF STEP 4
        
'Searches for when the value of the next cell is different than that of the current cell
'IF AAB DOES NOT EQUAL AAB (Rows 2 through 252)
'IF CELLS(253,1) does not equal CELLS(252,1)
'IF AAF does not equal AAB
'Then we do the following code:
        If Worksheets(cs).Cells(r + 1, 1).Value <> Worksheets(cs).Cells(r, 1).Value Then

'Step 2 - Script that loops through all stocks for single year and outputs yearly change from
'opening-closing price of that year
        change = Worksheets(cs).Cells(r, 6) - Worksheets(cs).Cells(start, 3)

'Step 3 - Script that loops through all stocks for single year and outputs percentage change from
'opening-closing price of that year
        percentage = change / Worksheets(cs).Cells(start, 3)
        start = r + 1
        
'Outputs
        Worksheets(cs).Range("K" & 2 + j).Value = percentage
        Worksheets(cs).Range("I" & 2 + j).Value = Worksheets(cs).Cells(r, 1).Value
        Worksheets(cs).Range("J" & 2 + j).Value = change
        Worksheets(cs).Range("K" & 2 + j).NumberFormat = "0.00%"
        Worksheets(cs).Range("Q2").NumberFormat = "0.00%"
        Worksheets(cs).Range("Q3").NumberFormat = "0.00%"
        Worksheets(cs).Range("L" & 2 + j).Value = totalvolume
        
'color coding to make column J cells be colored red for negative change or gree for positive change
        If change > 0 Then
            Worksheets(cs).Range("J" & 2 + j).Interior.ColorIndex = 4
        ElseIf change < 0 Then
            Worksheets(cs).Range("J" & 2 + j).Interior.ColorIndex = 3
        End If
        j = j + 1

'Step 5 - add functionality to the script to return the stock with "Greatest & increase", "Greatest % decrease"
'and "Greatest total volume"

If percentage > GreatestIncVal Then
                GreatestIncVal = percentage
                GreatestIncName = Worksheets(cs).Cells(r, 1).Value
            End If

If percentage < GreatestDecVal Then
                GreatestDecVal = percentage
                GreatestDecName = Worksheets(cs).Cells(r, 1).Value
            End If

If totalvolume > GreatestVolVal Then
                GreatestVolVal = totalvolume
                GreatestVolName = Worksheets(cs).Cells(r, 1).Value
            End If


'Step 6 - adjustments to VBA script to enable it to run on every worksheet (that is, every year) at once


totalvolume = 0

End If
Next r

Worksheets(cs).Range("P2").Value = GreatestIncName
Worksheets(cs).Range("Q2").Value = GreatestIncVal
Worksheets(cs).Range("P3").Value = GreatestDecName
Worksheets(cs).Range("Q3").Value = GreatestDecVal
Worksheets(cs).Range("P4").Value = GreatestVolName
Worksheets(cs).Range("Q4").Value = GreatestVolVal
GreatestIncName = ""
GreatestIncVal = 0
GreatestDecName = ""
GreatestDecVal = 0
GreatestVolName = ""
GreatestVolVal = 0

Next ws
End Sub
