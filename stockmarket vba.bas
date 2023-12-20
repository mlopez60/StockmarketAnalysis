Attribute VB_Name = "Module1"
Sub runinallsheets()
Dim sheet As Worksheet

For Each sheet In Worksheets
    sheet.Select
    Call stockmarket
Next

End Sub

Sub stockmarket()

'declare variables

Dim ticker As String
Dim openprice As Double
Dim closeprice As Double
Dim totalvolume As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim maxticker As String
Dim minticker As String
Dim totalticker As String

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Column = 1
totalvolume = 0
yearlychange = Range("C2").Value
percentchange = 0
Table = 2
maxincrease = 0
maxdecrease = 0
maxvolume = 0

' set up for loops

For i = 2 To lastrow


'Next cells comparison, if ticker symbols are different
If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then

'trackers
totalvolume = totalvolume + Cells(i, 7).Value
percentchange = (Cells(i, 6).Value - yearlychange) / yearlychange
yearlychange = Cells(i, 6).Value - yearlychange
ticker = Cells(i, Column).Value
    
    If percentchange > maxincrease Then
    maxincrease = percentchange
    maxticker = ticker
    ElseIf percentchange < maxdecrease Then
    maxdecrease = percentchange
    minticker = ticker
    End If
    
   If totalvolume > maxtotal Then
   maxtotal = totalvolume
   totalticker = ticker
   End If
    
'print ticker
Range("I1").Value = "Ticker"
Range("I" & Table).Value = ticker

'Print Yearly Change, format font color
Range("J1").Value = "Yearly Change"
Range("J" & Table).Value = yearlychange
If yearlychange < 0 Then
    Range("J" & Table).Interior.ColorIndex = 3
    Else
    Range("J" & Table).Interior.ColorIndex = 4
    
    End If

'Print Percent Change , format as percent
Range("K1").Value = "Percent Change"
Range("K" & Table).Value = FormatPercent(percentchange)

'Print volume
Range("L1").Value = "Total Stock Volume"
Range("L" & Table).Value = totalvolume
Table = Table + 1

'Print max
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Range("P2").Value = maxticker
Range("Q2").Value = FormatPercent(maxincrease)

Range("P3").Value = minticker
Range("Q3").Value = FormatPercent(maxdecrease)

Range("P4").Value = totalticker
Range("Q4").Value = maxtotal

'reset total volume for next ticker
totalvolume = 0
yearlychange = Cells(i + 1, 3).Value
percentchange = 0
'if ticker cells are the same, keep adding volume
Else
totalvolume = totalvolume + Cells(i, 7).Value
totalvolume = totalvolume


End If

Next i


End Sub

