Attribute VB_Name = "Module1"
Sub SummarizeStockData():

    'Initializing "i" and "j" counters for later
    'Initializing "currentStock" and "nextStock" for ticker comparison
    Dim ws As Worksheet
    Dim i, j, startIndex As Long
    Dim lastRow As Long
    Dim quarterlyDiff, percentDiff As Double
    Dim totalVolume As LongLong
    
    'Variables to hold biggest stats
    Dim biggestPercent, smallestPercent As Double
    Dim biggestVolume As LongLong

    For Each ws In Worksheets
        'Init variables to be reset with each worksheet
        Dim currentStock, nextStock As String
        totalVolume = 0   'Set beginning volume
        startIndex = 2      'Index where 1st instance of stock was found
        j = 2                     'j is the counter that holds the place for data printing after calculations
        
        'Begin Formatting of Sheets with Headers
        'Columns 9, 10, 11, 12 will store calculated data
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'Formatting 3x3 Grand Stats Section (Columns 15-17)
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        'Last row is calc per each worksheet
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To lastRow
            'Setting string vars to allow for easier writing of reoccuring vars
            currentStock = ws.Cells(i, 1).Value
            nextStock = ws.Cells(i + 1, 1).Value
            
            'Accruing total volume until new stock is found
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            'Checking to see if the next stock ticker is different than the current one
            'Main algorithm for calculations run only when condition is met
            If currentStock <> nextStock Then
                'Calculate the Quarterly and Quarterly % Differences
                quarterlyDiff = ws.Cells(i, 6).Value - ws.Cells(startIndex, 3).Value
                percentDiff = quarterlyDiff / ws.Cells(startIndex, 3).Value
                
                'Print the name & values into columns 9-12
                'Format Values as they're printing
                ws.Cells(j, 9).Value = currentStock
                ws.Cells(j, 10).Value = quarterlyDiff
                ws.Cells(j, 10).NumberFormat = "#0.00"  'Formatting for quarterly difference
                ws.Cells(j, 11).Value = percentDiff
                ws.Cells(j, 11).NumberFormat = "0.00%"  'Formatting for percent difference
                ws.Cells(j, 12).Value = totalVolume
                
                'Color coded quarterly difference conditional (red, green, or no format)
                If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = 0
                End If

                'Establish base set of biggest values for comparison when on 1st print entry
                'Prints base set to excel in case nothing exceeds the conditions specified
                If j = 2 Then
                    'Printing tickers
                    ws.Cells(2, 16).Value = currentStock
                    ws.Cells(3, 16).Value = currentStock
                    ws.Cells(4, 16).Value = currentStock
                    'Base set
                    biggestPercent = percentDiff
                    smallestPercent = percentDiff
                    biggestVolume = totalVolume
                    'Printing base set
                    ws.Cells(2, 17).Value = biggestPercent
                    ws.Cells(3, 17).Value = smallestPercent
                    ws.Cells(4, 17).Value = biggestVolume
                End If
                
                'Multiple solo "If" statements to check for multilateral differences
                'Trades and prints over values if conditions exceeded for record
                If biggestPercent < percentDiff Then
                    biggestPercent = percentDiff
                    ws.Cells(2, 16).Value = currentStock
                    ws.Cells(2, 17).Value = biggestPercent
                End If
                If smallestPercent > percentDiff Then
                    smallestPercent = percentDiff
                    ws.Cells(3, 16).Value = currentStock
                    ws.Cells(3, 17).Value = smallestPercent
                End If
                If biggestVolume < totalVolume Then
                    biggestVolume = totalVolume
                    ws.Cells(4, 16).Value = currentStock
                    ws.Cells(4, 17).Value = biggestVolume
                End If
                                
                'Increment "j" value for printing, reset volume, move to next stock and start index
                j = j + 1
                totalVolume = 0
                nextStock = currentStock
                startIndex = i + 1
            End If
        Next i
        
        'FINALIZE FORMATTING FOR NEW DATA - "j" should hold value of last row in printed table
        'Update % change column to percent style
        'After a Worksheet is done being built, autofit all columns
        ws.Range("K2:K" & j).NumberFormat = "0.00%"
        ws.Columns("A:Q").AutoFit
    Next ws

End Sub


