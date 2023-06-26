VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_market()
Dim ws As Worksheet
For Each ws In Worksheets
worksheetname = ws.Name
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

        
Dim TickCount As Long
Dim LastRowI As Long
Dim PerChange As Double
Dim GreatIncr As Double
Dim GreatDecr As Double
Dim GreatVol As Double
Dim i As Long
Dim j As Integer





TickCount = 2
 j = 2
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

        Dim TickerRow As Long: TickerRow = 1
Dim LastRowA As Long
For i = 2 To lastrow


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               TickerRow = TickerRow + 1
Ticker = ws.Cells(i, 1).Value
ws.Cells(TickerRow, "I").Value = Ticker
                
                
                
              
ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
ws.Cells(TickerRow, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
If ws.Cells(TickerRow, 10).Value < 0 Then
ws.Cells(TickerRow, 10).Interior.ColorIndex = 18
Else

ws.Cells(TickerRow, 10).Interior.ColorIndex = 31
        
 End If
 
 
 
 
 
 j = 2
Ticker = ws.Cells(i, 1).Value
If ws.Cells(j, 3).Value <> 0 Then
PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
 ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    
Else
PerChange = ((ws.Cells(j, 3).Value - ws.Cells(i, 6).Value) / ws.Cells(i, 6).Value)
ws.Cells(TickerRow, 11).Value = Format(PerChange, "Percent")
                    
End If

'ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))


'TickCount = TickCount + 1
 'j = i + 1
'j = 2

                End If
            
            Next i



 






Next ws


End Sub


