Attribute VB_Name = "Module4"

Sub AllStocksAnalysisRefactoredDynamic()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
       
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
       
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0
    Dim tickerIndexMax As Integer
    tickerIndexMax = 11
       
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    'looping through ticker data to read and store values for tickers()
    'have removed the static list
    For j = 2 To RowCount
    'check if this is the first row with this ticker - if so, add to array
    If Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
     tickers(tickerIndex) = Cells(j, 1).Value
     tickerIndex = tickerIndex + 1
     End If
    Next j
    
    'resetting tickerIndex to 0 for next loops
    tickerIndex = 0
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = tickerIndex To tickerIndexMax
     Ticker = tickers(tickerIndex)
     tickerVolumes(tickerIndex) = 0
        
    '2b) Loop over all the rows in the spreadsheet.
      For j = 2 To RowCount

        '3a) Increase volume for current ticker
        If Cells(j, 1).Value = Ticker Then
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(j, 1).Value = Ticker And Cells(j - 1, 1).Value <> Ticker Then
         tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(j, 1).Value = Ticker And Cells(j + 1, 1).Value <> Ticker Then
         tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
         '3d Increase the tickerIndex.
         tickerIndex = tickerIndex + 1
        End If
        
       Next j
       
     Next i
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
       
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

Sub thisisgoingtowork()
Dim tickers(12) As String
Dim tickerIndex As Integer
Worksheets("2017").Activate
tickerIndex = 0
For j = 2 To 3013
  'check if this is the first row with this ticker - if so, add to array
    If Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
          tickers(tickerIndex) = Cells(j, 1).Value
          tickerIndex = tickerIndex + 1
    End If
Next j
For i = 0 To 11
    Debug.Print tickers(i) 'prints first value(hardcoded) and then blank
Next i
End Sub
