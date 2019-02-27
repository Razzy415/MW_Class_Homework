Attribute VB_Name = "Module1"
Sub stockhomework()
'Loop through each worksheet
For Each ws In Worksheets

'Set variable for holding Ticker names
Dim ticker As String
'Set variable for holding the Total Stock Value per Ticker name
Dim totalstock As Double
totalstock = 0
'Declare variable for ticker row
Dim ticker_row As Double
'Track ticker row
ticker_row = 2
'Declare variable for stock opening value
Dim stock_open As Double
'Declare variable for stock closing value
Dim closevalue As Double
'Initialize and check open and close value
check_open = ""
check_close = ""
Dim gincrease As Double

'Determine the last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Insert Column for Yearly Change and Percet Change
Columns("J:K").Insert Shift:=xlToRight

'Create Headers for Ticker, Yearly Change, Percent Change, Total Stock Volume, etc
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Value"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Range("I1").EntireColumn.AutoFit
    ws.Range("J1").EntireColumn.AutoFit
    ws.Range("K1").EntireColumn.AutoFit
    ws.Range("L1").EntireColumn.AutoFit
    ws.Range("O1").EntireColumn.AutoFit
    ws.Range("P1").EntireColumn.AutoFit
    ws.Range("Q1").EntireColumn.AutoFit

      ' Loop through all Ticker names
  For i = 2 To LastRow

    ' Ticker name check
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      ' Set the Ticker name
      ticker = ws.Cells(i, 1).Value
      
      'Track and check close value
      stockclose = ws.Range("F" & i)
      check_close = "y"
      
      'Print yearly change and percent change while check for any occurences of null value
      If stock_open <> 0 Then
            ws.Range("J" & ticker_row).Value = stockclose - stock_open
            ws.Range("K" & ticker_row).Value = (stockclose - stock_open) / stock_open
           Else
            ws.Range("J" & ticker_row).Value = stockclose - stock_open
            ws.Range("K" & ticker_row).Value = 0
           End If

      'Format percentage
      ws.Range("K" & ticker_row).NumberFormat = "0.00%"
      ' Add to the totalstock
      totalstock = totalstock + ws.Cells(i, 7).Value

      ' Print tickername
      ws.Range("I" & ticker_row).Value = ticker

      ' Print the totalstock
      ws.Range("L" & ticker_row).Value = totalstock
 
      'Conditiaonals and color cells
        If ws.Range("J" & ticker_row) >= 0 Then
          ws.Range("J" & ticker_row).Interior.ColorIndex = 4
         ElseIf ws.Range("J" & ticker_row) < 0 Then
           ws.Range("J" & ticker_row).Interior.ColorIndex = 3
         End If

      ' Add one to the ticker_row
      ticker_row = ticker_row + 1
      
      ' Reset the totalstock for the next ticker name
      totalstock = 0

      gincrease = 0
     
   'Grab opening stock and put a stop
    Else
      If check_open = "" Then
          stock_open = ws.Range("C" & i)
          check_open = "y"
      End If

      ' Add to the totalstock
      totalstock = totalstock + ws.Cells(i, 7).Value

End If
        
Next i

Next ws

End Sub


