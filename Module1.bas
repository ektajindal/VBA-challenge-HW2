Attribute VB_Name = "Module1"
Sub Multiple_year_stock()

'Declare Worksheet

Dim ws As Worksheet

'Loop through all stcoks for one year

For Each ws In Worksheets

'Create column headings

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest Per Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

Dim Ticker As String
Dim Ticker_Volume As Double
Ticker_Volume = 0
Dim open_value As Double
Dim close_value As Double
Dim yearly_change As Double
Dim percent_change As Double
yearly_change = 0

Dim i As Long
Dim TickerRow As Integer
TickerRow = 2
Dim Lastrow As Long
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set the open value

open_value = ws.Cells(TickerRow, 3).Value


For i = 2 To Lastrow

'Check if Ticker name changed

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Set the Ticker Name
Ticker = ws.Cells(i, 1).Value

'print the Ticket Name in the Summary

ws.Range("I" & TickerRow).Value = Ticker

'calculate yearly change
yearly_change = yearly_change + (ws.Cells(i, 6).Value - open_value)
ws.Range("J" & TickerRow).Value = yearly_change

'calculate percent change
percent_change = (yearly_change / open_value)
ws.Range("K" & TickerRow).Value = percent_change
ws.Range("K" & TickerRow).Style = "Percent"
ws.Range("K" & TickerRow).NumberFormat = "0.00%"

'Add the ticker volume
Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
     
        
'Print Ticker Volumne to the Summary
ws.Range("L" & TickerRow).Value = Ticker_Volume
 
 'Add one to the ticker row
 TickerRow = TickerRow + 1

'Reset the total volume total
Ticker_Volume = 0
yearly_change = 0
open_value = ws.Cells(i + 1, 3).Value


Else

'Add the Volume Total

Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

End If

 Next i
 
 'Declare variable for formatting
 
 Dim endrow As Long
 
 endrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
 
 'Add Loop for Formatting
 
 For i = 2 To endrow
 
 If ws.Cells(i, 11).Value >= 0 Then
 
  ws.Cells(i, 11).Interior.ColorIndex = 4
 
 Else
 
   ws.Cells(i, 11).Interior.ColorIndex = 3
    
End If

Next i

' Declare variable for greatest total volume

Dim totalVolumeRow As Long

totalVolumeRow = ws.Cells(Rows.Count, 12).End(xlUp).Row

Dim totalVolumeMax As Double

totalVolumeMax = 0

'Add Loop for finding greatest total volume
For i = 2 To totalVolumeRow

'Add Conditional for greatest total volume

    If totalVolumeMax < ws.Cells(i, 12).Value Then
        totalVolumeMax = ws.Cells(i, 12).Value
        ws.Cells(4, 17).Value = totalVolumeMax
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    End If
Next i

'Declare variables for finding Greatest % Increase and Greatest % decrease

Dim percentLastRow As Long
percentLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
Dim percent_max As Double
percent_max = 0
Dim percent_min As Double
percent_min = 0

'Add Loop

For i = 2 To percentLastRow

'Add Conditions

    If percent_max < ws.Cells(i, 11).Value Then
        percent_max = ws.Cells(i, 11).Value
        ws.Cells(2, 17).Value = percent_max
        ws.Cells(2, 17).Style = "Percent"
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        
    ElseIf percent_min > ws.Cells(i, 11).Value Then
        percent_min = ws.Cells(i, 11).Value
        ws.Cells(3, 17).Value = percent_min
        ws.Cells(3, 17).Style = "Percent"
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        
    End If
Next i

Next ws


End Sub
