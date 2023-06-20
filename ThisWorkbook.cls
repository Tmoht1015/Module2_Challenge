VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Mult_Yr_Stock_Data()

' Variables
For Each ws In Worksheets

 Dim wsname As String
 Dim x As Long
 Dim y As Long
Dim LastRowTicker As Long
 Dim TickerCount As Long
 Dim PercentChg As Double
 Dim GreatestInc As Double
 Dim GreatestDec As Double
 Dim GreatestVolume As Double
 
 wsname = ws.Name
 
 ' Headers
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Annual Change"
 ws.Cells(1, 11).Value = "% Change"
 ws.Cells(1, 12).Value = "Total Volume"
 ws.Cells(1, 16).Value = "Ticker"
 ws.Cells(1, 17).Value = "Value"
 ws.Cells(2, 15).Value = "Greatest % Increase"
 ws.Cells(3, 15).Value = "Greatest % Decrease"
 ws.Cells(4, 15).Value = "Greatest Total Vol."
 
 ' Ticker Counter Set
 TickerCount = 2
 y = 2
 
 ' Looping
  For x = 2 To 753001
  
  ' Ticker Check
  If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
  ws.Cells(TickerCount, 9).Value = ws.Cells(x, 1).Value
  ws.Cells(TickerCount, 10).Value = ws.Cells(x, 6).Value - ws.Cells(y, 3).Value
 
 ' Colors
 If ws.Cells(TickerCount, 10).Value < 0 Then
 
 ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
 
 Else
 
 ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
 
 End If
 
 ' % Change Calculation
 If ws.Cells(y, 3).Value <> 0 Then
 
 PercentChg = ((ws.Cells(x, 6).Value - ws.Cells(y, 3).Value) / ws.Cells(y, 3).Value)
 
 ws.Cells(TickerCount, 11).Value = Format(PercentChg, "Percent")
 
 Else
 ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
 End If
 
 ' Total Volume
 ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(y, 7), ws.Cells(x, 7)))
 
 TickerCount = TickerCount + 1
 
 y = x + 1
 
 End If
 
 Next x
 
 ' Summary_Prep
LastRowTicker = ws.Cells(Rows.Count, 9).End(xlUp).Row

 GreatestInc = ws.Cells(2, 11).Value
 GreatestDec = ws.Cells(2, 11).Value
 GreatestVolume = ws.Cells(2, 12).Value
 
 ' Summary_Loops_TickerColumn
 For x = 2 To LastRowTicker
 
 If ws.Cells(x, 11).Value > GreatestInc Then
 GreatestInc = ws.Cells(x, 11).Value
 ws.Cells(2, 16).Value = ws.Cells(x, 9).Value
 
 Else
 
 GreatestInc = GreatestInc
 
 End If
 
 If ws.Cells(x, 11).Value < GreatestDec Then
 GreatestDec = ws.Cells(x, 11).Value
 ws.Cells(3, 16).Value = ws.Cells(x, 9).Value
 
 Else
 
 GreatestDec = GreatestDec
 
 End If
 
 If ws.Cells(x, 12).Value > GreatestVolume Then
 GreatestVolume = ws.Cells(x, 12).Value
 ws.Cells(4, 16).Value = ws.Cells(x, 9).Value
 
 Else
 
 GreatestVolume = GreatestVolume
 
 End If
 
 ' Summary_Print
 ws.Cells(2, 17).Value = Format(GreatestInc, "Percent")
 ws.Cells(3, 17).Value = Format(GreatestDec, "Percent")
 ws.Cells(4, 17).Value = Format(GreatestVolume, "Scientific")
  
  Next x
  
  Next ws
 End Sub

