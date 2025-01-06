Attribute VB_Name = "VBA_challenge"
Sub VBA_challenge()
'Create worksheet variable'
   Dim ws As Worksheet
'Start loop ro loop through all worksheets'
   For Each ws In Worksheets
'Create variables'
    Dim currentRow As Long, lastrow As Long
    Dim ticker As String
    Dim summaryRow As Integer
    Dim volumeTotal As Double, percentageChange As Double, quarterlyChange As Double
    Dim openStart As Double, closeEnd As Double
'Initialize variable values'
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    summaryRow = 2
    openStart = ws.Cells(2, 3).Value
    volumeTotal = 0
'Write headers for summary cells'
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
'Start for loop through all rows'
    For currentRow = 2 To lastrow
    'Set conditions for a change in Ticker symbol'
        If ws.Cells(currentRow + 1, 1).Value <> ws.Cells(currentRow, 1).Value Then
            'Grab ticker, volumeTotal, ending close price, and calculated quarterly and percentage change'
            ticker = ws.Cells(currentRow, 1).Value
            volumeTotal = volumeTotal + ws.Cells(currentRow, 7).Value
            closeEnd = ws.Cells(currentRow, 6).Value
            quarterlyChange = closeEnd - openStart
            percentageChange = (closeEnd - openStart) / openStart
            'Write in summary values for ticker, quarterly and percentage change'
            ws.Range("I" & summaryRow).Value = ticker
            ws.Range("L" & summaryRow).Value = volumeTotal
            ws.Range("J" & summaryRow).Value = quarterlyChange
                'Set conditional formatting for quarterly and percentage change'
                If quarterlyChange > 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 2
                End If
            ws.Range("K" & summaryRow).Value = percentageChange
            ws.Range("K" & summaryRow).NumberFormat = "0.00%"
                If percentageChange > 0 Then
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 4
                ElseIf percentageChange < 0 Then
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(summaryRow, 11).Interior.ColorIndex = 2
                End If
            'Move summary row and reset volume total and openStart values'
            summaryRow = summaryRow + 1
            volumeTotal = 0
            openStart = ws.Cells(currentRow + 1, 3).Value
    
        Else
            'Add volume total if ticker is the same'
            volumeTotal = volumeTotal + ws.Cells(currentRow, 7).Value
        End If
        
    Next currentRow
   'Write header values for another summary'
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
   'Create variables'
    Dim currentRow2 As Long, lastRow2 As Long
    Dim maxTicker As String, minTicker As String, volTicker As String
    Dim maxPercent As Double, minPercent As Double, maxVolume As Double
    'Set initial values'
    maxPercent = 0
    minPercent = 0
    maxVolume = 0
    lastRow2 = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
   'Start for loop through summary rows'
    For currentRow2 = 2 To lastRow2
        'Create conditional checking if current percent change is greater than max change or less than min change'
     If ws.Cells(currentRow2, 11).Value > maxPercent Then
        maxPercent = ws.Cells(currentRow2, 11).Value
        maxTicker = ws.Cells(currentRow2, 9).Value
     End If
     If ws.Cells(currentRow2, 11).Value < minPercent Then
        minPercent = ws.Cells(currentRow2, 11).Value
        minTicker = ws.Cells(currentRow2, 9).Value
     End If
        'Create conditional checking if current volume is greater than max volume'
     If ws.Cells(currentRow2, 12).Value > maxVolume Then
        maxVolume = ws.Cells(currentRow2, 12).Value
        volTicker = ws.Cells(currentRow2, 9).Value
     End If
    Next currentRow2
   'Write in the stored values for tickers, percentage changes, and max volume'
     ws.Range("P2").Value = maxTicker
     ws.Range("P3").Value = minTicker
     ws.Range("Q2").Value = maxPercent
     ws.Range("Q3").Value = minPercent
     ws.Range("P4").Value = volTicker
     ws.Range("Q4").Value = maxVolume
     ws.Range("Q2").NumberFormat = "0.00%"
     ws.Range("Q3").NumberFormat = "0.00%"
'Continue to next worksheet'
    Next ws
   
End Sub
