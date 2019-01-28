Sub StockExchange()

'for that selects each worksheet (separately) and calls the countrow, MinMax, and format subs
'call Summary()
'call MinMax()
'call format()
'next worksheet
Dim CtWkst As Integer
CtWkst = ThisWorkbook.Worksheets.Count - 1
For i = 1 To CtWkst
Worksheets(i).Activate
Call Summary
Call MinMax
Call Fortmatting
Next i
Debug.Print CtWkst
End Sub
'
' Subroutine Summary
'
Sub Summary()
'summarizes original/given data
Dim RCt As Long
Dim RwSt As Integer
Dim TRow As Integer
Dim TickerName As String
Dim TStockVol As Long
Dim TOpen As Double
Dim TClose As Double
Dim TYearly As Double
Dim PercentChange As Double

'sets inital values.  counts number of rows with data.  defines starting row of numbers (RwSt), TRow specfies rows for the summaries
TYearly = 0
PercentChange = 0
TStockVol = 0
RCt = Cells(Rows.Count, 7).End(xlUp).Row
TRow = 2
RwSt = 2
'while sums values with same ticker name
For i = RwSt To RCt
  TickerName = Cells(i, 1)
  Do While Cells(i, 1).Value = Cells(i + 1, 1).Value
     TOpen = TOpen + Cells(i, 3).Value
     TClose = TClose + Cells(i, 6).Value
     TStockVol = TStockVol + Cells(i, 7).Value / 100
     i = i + 1
  Loop
'category summaries
  TOpen = TOpen + Cells(i, 3).Value
  TClose = TClose + Cells(i, 6).Value
  TYearly = TClose - TOpen
   Debug.Print (i)
  'if statement for the tricky Ticker with all zeros.
  If TOpen = 0 Then
     PercentageChange = 0
     Else
  PercentChange = (TClose - TOpen) / TOpen
  End If
  'Debug.Print (PercentChange)
  'Debug.Print (TOpen)
  'Debug.Print (TClose)
 'TStockVol = TStockVol + Cells(i, 7).Value / 100
  TStockVol = TStockVol + Cells(i, 7).Value
'Places summary values into summary 'table'
  Cells(TRow, 9).Value = TickerName
  Cells(TRow, 10).Value = TYearly
  Cells(TRow, 11).Value = PercentChange
  Cells(TRow, 12).Value = TStockVol

'Resets variables. Increases Row where new ticketer summaries will be placed
  TRow = TRow + 1
  TYearly = 0
  TOpen = 0
  TClose = 0
  PercentChange = 0
  TStockVol = 0

Next i
End Sub
'
' Subroutine MinMax 
'
Sub MinMax()
'Second round of calulations for the "hard" section.  Evaluates Min/Max values from the summary tables
Dim MinTicker As String
Dim MaxTicker As String
Dim MaxVolTicker As String
Dim MinPercentage As Double
Dim MaxPercentage As Double
Dim MaxStock As Long
Dim HCt As Integer
Dim RwSt As Integer

'counts number of rows in summary table and sets row data starts
HCt = Cells(Rows.Count, 9).End(xlUp).Row
RwSt = 2
'sets initial values of variables
MinPercentage = Cells(RwSt, 11).Value
MaxPercentage = Cells(RwSt, 11).Value
MaxStock = Cells(RwSt, 12).Value
MinTicker = Cells(RwSt, 9).Value
MaxTicker = Cells(RwSt, 9).Value
MaxVolTicker = Cells(RwSt, 9).Value

'For loop finds min/max and grabs the associated tickername.  Will cycle through all rows in summary table.
For H = RwSt To HCt
  If MinPercentage > Cells(H + 1, 11).Value Then
    MinPercentage = Cells(H + 1, 11).Value
    MinTicker = Cells(H + 1, 9).Value
    End If
  If MaxPercentage < Cells(H + 1, 11).Value Then
    MaxPercentage = Cells(H + 1, 11).Value
    MaxTicker = Cells(H + 1, 9).Value
    End If
  If MaxStock < Cells(H + 1, 12).Value Then
    MaxStock = Cells(H + 1, 12).Value
    MaxVolTicker = Cells(H + 1, 9).Value
    End If
Next H

'Places the max/mins and associated tickers into "hard" table
Cells(2, 16).Value = MaxTicker
Cells(3, 16).Value = MinTicker
Cells(4, 16).Value = MaxVolTicker
Cells(2, 17).Value = MaxPercentage
Cells(3, 17).Value = MinPercentage
Cells(4, 17).Value = MaxStock

End Sub
'
' Subroutine Formatting 
'

Sub Fortmatting()

Dim RwSt As Integer
Dim RCt As Long

RwSt = 2
RCt = Cells(Rows.Count, 7).End(xlUp).Row


'format conditional if = 0, grey; if >0, green; else red. used colorchart http://dmcritchie.mvps.org
For i = RwSt To RCt
  If Cells(i, 10).Value = 0 Then
    Cells(i, 10).Interior.ColorIndex = 34
  ElseIf Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 43
  Else
    Cells(i, 10).Interior.ColorIndex = 3
  End If
Next i
     

End Sub

