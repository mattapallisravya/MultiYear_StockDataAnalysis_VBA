Attribute VB_Name = "Module1"
Public Sub WallStreetAnalysis()
Dim ws As Worksheet
'loop to run on everysheet
For Each ws In Worksheets
ws.Activate
ActiveSheet.UsedRange.EntireColumn.AutoFit
Dim Lastrow As Long


Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
TotalVolume = 0
SummaryCounter = 2
OpeningPrice = Cells(2, 3)
GreatestPercentageIncrease = 0
GreatestPercentageIncreaseTicker = ""
GreatestPercentageDecrease = 0
GreatestPercentageDecreaseTicker = ""
GreatestTotalVolume = 0
GreatestTotalVolumeTicker = ""

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

For i = 2 To Lastrow

ClosingPrice = Cells(i, 6).Value

If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
    'Calculating Yearly Change,Percentage Change, Total Stock Volume and corresponding Ticker
   
     Cells(SummaryCounter, 9).Value = Cells(i, 1).Value
     Cells(SummaryCounter, 10).Value = ClosingPrice - OpeningPrice
     Cells(SummaryCounter, 11).Value = Cells(SummaryCounter, 10).Value / OpeningPrice
     Cells(SummaryCounter, 11).NumberFormat = "0.00%"
     
     'Greatest Percentage Increase
     If (Cells(SummaryCounter, 11).Value > GreatestPercentageIncrease) Then
     GreatestPercentageIncrease = Cells(SummaryCounter, 11).Value
     GreatestPercentageIncreaseTicker = Cells(SummaryCounter, 9).Value
     End If
     
     'Greatest Percentage Decrease
     If (Cells(SummaryCounter, 11).Value < GreatestPercentageDecrease) Then
     GreatestPercentageDecrease = Cells(SummaryCounter, 11).Value
     GreatestPercentageDecreaseTicker = Cells(SummaryCounter, 9).Value
     End If
     
     'Greatest Total Volume
     If (Cells(SummaryCounter, 12).Value > GreatestTotalVolume) Then
     GreatestTotalVolume = Cells(SummaryCounter, 12).Value
     GreatestTotalVolumeTicker = Cells(SummaryCounter, 9).Value
     End If
     
     'Conditional Formating
     If Cells(SummaryCounter, 10).Value = 0 Then
     Cells(SummaryCounter, 10).Interior.ColorIndex = 2
     ElseIf Cells(SummaryCounter, 10).Value > 0 Then
     Cells(SummaryCounter, 10).Interior.ColorIndex = 4
     Else
     Cells(SummaryCounter, 10).Interior.ColorIndex = 3
     End If
     Cells(SummaryCounter, 12).Value = TotalVolume + Cells(i, 7)
    
     TotalVolume = 0
     OpeningPrice = Cells(i + 1, 3)
     SummaryCounter = SummaryCounter + 1
    
   Else
TotalVolume = TotalVolume + Cells(i, 7)

End If

Next i
'Assigning Values to corresponding Cells
Cells(2, 16).Value = GreatestPercentageIncrease
Cells(2, 16).NumberFormat = "0.00%"
Cells(2, 15) = GreatestPercentageIncreaseTicker
Cells(3, 16).Value = GreatestPercentageDecrease
Cells(3, 16).NumberFormat = "0.00%"
Cells(3, 15) = GreatestPercentageDecreaseTicker
Cells(4, 15) = GreatestTotalVolumeTicker
Cells(4, 16) = GreatestTotalVolume

Next ws

End Sub

