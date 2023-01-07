Sub StockAnalysis()

  'Get ready to loop through each sheet in the workbook
  For Each ws In Worksheets
  

  ' Set a variable for specifying the column of interest
  Dim column As Integer
  Dim Ticker As String
  Dim YearlyChange As Double
  Dim StockVolume As Double

  
  'Set up the two output tables with headers/rows
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"
  
  ws.Cells(2, 15).Value = "Greatest % Increase"
  ws.Cells(3, 15).Value = "Greatest % Decrease"
  ws.Cells(4, 15).Value = "Greatest Total Volume"
  
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"
  
  
  'Find Beginning Cell for each sequence to define the range
  StartofRange = 2

  'Set the start of the output table
  Output_Table_Row = 2
  
  'Define the column we are looping through
  column = 1

  ' Loop through each row in the first column and find the last row
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  For i = 2 To lastrow
  
  
  'Add up the stock volume for a symbol
  StockVolume = StockVolume + ws.Cells(i, 7).Value
  TotalCellsforTicker = TotalCellsforTicker + ws.Cells(i, 7).Value
  

    ' Searches for when the value of the next cell is different than that of the current cell
    ' THIS HAPPENS PRIOR TO NEXT TICKER SYMBOL
    If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
    
    'Set a variable to find the end cell in the range
    EndofRange = i
    
      Ticker = ws.Cells(i, column).Value
      ws.Cells(Output_Table_Row, 9).Value = Ticker
      ws.Cells(Output_Table_Row, 12).Value = StockVolume
      
      
       'Calculate annual change in stock - day 1 open - last day close
       YearlyChange = ws.Cells(EndofRange, 6).Value - ws.Cells(StartofRange, 3).Value
       ws.Cells(Output_Table_Row, 10).Value = YearlyChange
       
       If YearlyChange < 0 Then
       ws.Cells(Output_Table_Row, 10).Interior.ColorIndex = 3
       
       Else
       ws.Cells(Output_Table_Row, 10).Interior.ColorIndex = 4
       
       End If
       
       'Calculate the percentage change and format cell as a percent
       StockPercentage = YearlyChange / ws.Cells(StartofRange, 3).Value
       ws.Cells(Output_Table_Row, 11).Value = FormatPercent(StockPercentage)
       
      
      'increment rows in output table
      Output_Table_Row = Output_Table_Row + 1
      
      
      ' Reset StockVolume before next symbol
      StockVolume = 0
      
      'Reset start of range
      StartofRange = i + 1

    End If

  Next i
  
  'Second verse, same as the first. But this time, loop the output table
  column = 11

  'Find the min/max values in the range for the ticker selected
  Rng = "K:K"
  MaxStock = Application.WorksheetFunction.Max(ws.Range(Rng))
  MinStock = Application.WorksheetFunction.Min(ws.Range(Rng))
  
  'Find the max volume of trading based on the ticker selected
  MaxVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
  
    ' Loop through each row in the first column and find the last row
  
  'Find the last roo in the spreadsheet
  lastrow = ws.Cells(Rows.Count, column).End(xlUp).Row
  For i = 2 To lastrow
  
    'Populate the output table
    If ws.Cells(i, column).Value = MaxStock Then
    ws.Cells(2, 16).Value = ws.Cells(i, 9)
    ws.Cells(2, 17).Value = FormatPercent(MaxStock)
    
    ElseIf ws.Cells(i, column).Value = MinStock Then
    ws.Cells(3, 16).Value = ws.Cells(i, 9)
    ws.Cells(3, 17).Value = FormatPercent(MinStock)
    
    ElseIf ws.Cells(i, 12).Value = MaxVolume Then
    ws.Cells(4, 16).Value = ws.Cells(i, 9)
    ws.Cells(4, 17).Value = MaxVolume
    End If
    
    'Adjust all the column widths for readability
    ws.Columns("I:Q").AutoFit
  
  Next i
    
  Next ws

End Sub

