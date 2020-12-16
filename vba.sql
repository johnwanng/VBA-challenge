Sub Vba_Homework()
' Define summary array for the Greatest summary tickers
Dim tickerArray(2) As String

' Define summary array for the Greatest summary total
Dim greatestArray(2) As Double

'greatestArray(0) holds the Greatest Increase %
'tickerArray(0) holds the ticker having the Greatest Increase %

'greatestArray(1) holds the Greatest Decrease %
'tickerArray(1) holds the ticker having the Greatest Decrease %

'greatestArray(2) holds the Greatest Volume
'tickerArray(2) holds the ticker having the Greatest Volume

' Define currentTicker
Dim currentTicker As String
Dim yearlyChange As Double
Dim percentageChange As Double
Dim stockVolume As Double
Dim openingValue As Double
Dim closingValue As Double
Dim lastRow As Long
Dim summaryRow As Long
' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------

For Each ws In Worksheets

    'Initialise summaryRow to 1
    summaryRow = 1

    'Prepare the summary column headings for the given work sheet
    ws.Cells(summaryRow, 9) = "Ticker"
    ws.Cells(summaryRow, 10) = "Yearly Change"
    ws.Cells(summaryRow, 11) = "Percentage Change"
    ws.Cells(summaryRow, 12) = "Total Stock Volume"
    
    ' Initialise summary array variables for every work sheet
    For k = 0 To 2
        tickerArray(k) = ""
        greatestArray(k) = 0
    Next k
    
    currentTicker = ""
    openingValue = 0
    percentageChange = 0
    stockVolume = 0
    closingValue = 0
    yearlyChange = 0
    ' Find the last row for the current work sheet
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Start from row 2 and all the way to the last row for that work sheet
    For i = 2 To lastRow
          ' Assign current ticker and openingValue for the first row of the ticker
          If currentTicker = "" Then
            'Assign the current ticker in column A
            currentTicker = ws.Cells(i, 1).Value
            'Assign the openingValue in column C
            openingValue = ws.Cells(i, 3).Value
          End If
          'Increment the stockVolume from column G
          stockVolume = stockVolume + ws.Cells(i, 7)
          ' When the value of the next ticker is different than the current one
          If ws.Cells(i + 1, 1).Value <> currentTicker Then
                      ' Increment the summaryRow as it is time to display summary information
                      summaryRow = summaryRow + 1
                      'Store the closing value held in column F in the last row of the current ticker
                      closingValue = ws.Cells(i, 6)
                      ' Store the yearlyChange between closing value minus openingValue
                      yearlyChange = closingValue - openingValue
                      ws.Cells(summaryRow, 9) = currentTicker
                      ws.Cells(summaryRow, 10) = yearlyChange
                     ' Set the Cell Colors to Green if positive or zero otherwise set it to red
                      If yearlyChange >= 0 Then
                        ' Set the Cell Colors to Green
                        ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                      Else
                        ' Set the Cell Colors to Red
                        ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                      End If
                      'If either yearlyChange = 0 Or openingValue = 0 then
                      'set the Percentage = 0 to avoid overflow or divided by 0 errors
                      If yearlyChange = 0 Or openingValue = 0 Then
                         percentageChange = 0
                         ws.Cells(summaryRow, 11) = 0
                      Else
                         percentageChange = yearlyChange / openingValue
                         ws.Cells(summaryRow, 11) = percentageChange
                      End If
                      'Set cell to have percentage format and style
                      ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                      ws.Cells(summaryRow, 12) = stockVolume
                      
                      'Compare the value in greatestArray with current percentageChange and replace them according their specific criteria
                      'tickerArray value will get replaced if their greatestArray criteria is met as well
                      If greatestArray(0) < percentageChange Then
                         tickerArray(0) = currentTicker
                         greatestArray(0) = percentageChange
                      End If
                      If greatestArray(1) > percentageChange Then
                         tickerArray(1) = currentTicker
                         greatestArray(1) = percentageChange
                      End If
                      If greatestArray(2) < stockVolume Then
                         tickerArray(2) = currentTicker
                         greatestArray(2) = stockVolume
                      End If
                      ' Initialise those tally variables for the next ticker
                      currentTicker = ""
                      yearlyChange = 0
                      stockVolume = 0
                      openingValue = 0
                      closingValue = 0
                      percentageChange = 0
          End If
    Next i
    'Display Summary titles and information
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    'Display Summary information from the tickerArray and greatestArray
    For x = 0 To 2
        ws.Cells(2 + x, 15).Value = tickerArray(x)
        ws.Cells(2 + x, 16).Value = greatestArray(x)
        'Only set to percentage format and style for the first 2 array index, i.e. Greatest % Increase and Greatest % Decrease
        If x < 2 Then
            'Set cell to have percentage format and style
            ws.Cells(2 + x, 16).NumberFormat = "0.00%"
        End If
    Next x
Next ws

        

End Sub