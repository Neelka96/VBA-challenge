Attribute VB_Name = "Module1"
'Version: 2.0
'    ---------------------------------
' /   Written by: Neel Agarwal @Neelka96    \
'/  Last Updated: (mm.dd.yyyy) 12.2.2024  \
'--------------------------------------
'MODULE PURPOSE:
'      Module 's purpose is to analyze a large amount of data held on multiple pages of a workbook. The data consists
'      of daily stock values and the objective is to provide a summary for each one on each sheet, along with a smaller
'      summary for each worksheet. After an examination of the data, the module creates new tables of data with
'      corresponding formatting and headers that will be filled in with computations that have been run on each grouping
'      of stock tickers to create the summaries. The workbook and any worksheets within it should only update after the
'      module finishes its execution along with message boxes declaring the macro a success and lisitng the execution time.
'
'CURRENT METHOD PATH:
'      <Run Macro>
'      --> Disable application update features --> Begin loop through each worksheet
'      --> Copies data and perserves formatting/arrangement in same 2D array dimensions
'      --> Runs data through tests in local memory --> Stores results in two new 2D arrays
'      --> Two new arrays with headings and formatting printed back to workshees
'      --> Loop to next worksheet
'      --> Success! Re-enable application update features to see changes
'      --> Print message box stating success and execution time
'      <End Macro>
'
'DETAILED DESCRIPTION OF MODULE:
'      Built to work with daily entries of various stocks and their associated values listed in one or more worksheets. Suboutine
'      takes the input of a workbook and copies pre-organized data within the workbook (and its dimensions) to the temporary
'      memory of local computer in the form of a variant array. The array data include the stock name, date of entry, the opening,
'      high, low, and closing prices, and volume for each entry. The sample data given by the class has the workbook seperated into
'      quarters and the stocks grouped by tickers which makes comparing groups of stocks easier. Looping within the new array,
'      it only pauses at new stock tickers to compute the statistics of the prior range of stock tickers. The module will summarize
'      the data for each stock ticker group and working one quarter at a time until the output is completed. The computations are
'      stored in two new 2D arrays on computer by grouping of stock ticker. After new arrays are built, data is printed back to each
'      worksheet with macro-based conditional formatting. Finally, the module permits the Microsoft Excel UI to resume updates
'      before subroutine end. The current subroutine is now built using arrays to handle large quantities of data with rapid efficiency!
'      The previous subroutine ran by continuously accessing Microsoft Excel Ranges in the workbook and storing all temporary data
'      within the sheet, however it executed at only 20% of the speed of current mode.
'
'CURRENT LIMITATIONS:
'      Runs only with Microsoft Excel Workbooks.
'      Stock tickers don't have to be in alphabetical order, but they MUST be grouped by ticker name.
'      Stock tickers in the same group MUST be in exact chronological order (Earliest date --> Latest Date).
'      Input format must be: {Rows} x {Columns} --> {Ticker, date, open, high, low close, volume} x {Daily stock entries}
'      Input data is NOT verified/formatted/cleansed by module - adding said functionality has not been addressed.
'-----------------------------------------------------------------------------------------
Sub vbaSummarizeStocks():
    '(-----------------------------------------)
    '(                CHECKING FOR SPEED HERE                 )
    '(-----------------------------------------)
    Dim startTime, endTime, elapsedTime As Single   ')
    startTime = Timer                                                ')
    '(-----------------------------------------)
    '######################################################################
    '# Title: ChatGPT's suggested use of Application.<Property> temporarily turns off Excel
    '#       Worksheet properties that run during macro use to improve latency
    '# Author: OpenAI
    '# Date: 2024
    '# Code Version: ChatGPT-4o
    '# Availability: https://www.chatgpt.com
    '######################################################################
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '------------------------------------
    'DECLARING VARIABLES IN GLOBAL SCOPE
    '------------------------------------
    Dim ws As Worksheet   'Worksheet looping variable

    '##################################################################################
    '# Title: ChatGPT suggested use of RAM (variant type array method) for storing, accessing, computing,
    '#       comparing, and printing data. Implemented for its speed. Algorithm still implemented by @Neelka96
    '# Author: OpenAI
    '# Date: 2024
    '# Code Version: ChatGPT-4o
    '# Availability: https://www.chatgpt.com
    '###################################################################################
    Dim stockInput As Variant   'Array will hold all original data (aka Input)
    'Has dimensions of: {# Columns} x {# Rows} or more precisely...
    '{Ticker, date, open, high, low close, volume} x {Distinct stock entry dates}
    
    Dim stockOutput() As Variant   'Array will hold all output values - ReDim'd later at exact sizing
    'Has dimensions of: {# Columns} x {# Rows} or more precisely...
    '{Ticker, quarterly difference, quarterly percent difference, total volume} x {Distinct stock tickers}
    
    Dim stockStats(1 To 3, 1 To 2) As Variant   'Array will hold largest values within stockOutput() for output also
    'Has dimensions of: {# Columns} x {# Rows} or more precisely...
    '{Ticker, value} x {Greatest % increase, greatest % decrease, greatest volume}
    
    '-----------------
    'BEGIN WS LOOPING
    '-----------------
    For Each ws In Worksheets
        Dim i As Long   'Universal counter variable used in multiple "for" loops throughout code
        
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row   'Last row calc per ws to initialize stockInput
        stockInput = ws.Range("A2:G" & lastRow).Value   'ALL ORIGINAL DATA STORED IN ARRAY
        
        'Determine necessary size of stockOutput
        '-----------------------------------
        Dim outputCount As Long
        outputCount = 1   'Sets inital count of print entries to 1
        '#################################################################################
        '# Title: Information on UBound(), With, ReDim declaration, variant data type, and conditional formatting
        '# Author: Microsoft Learn/o365devx/AlexJerabek/kbrandl/OfficeGSX/Saisang
        '# Date: 2021-2022
        '# Availability: https://learn.microsoft.com/en-us/office/vba/api/overview/
        '#################################################################################
        For i = 2 To UBound(stockInput, 1)   'Looping from 2nd index to last index of stockInput()
            If stockInput(i, 1) <> stockInput(i - 1, 1) Then   'Checks if prior value is different
                outputCount = outputCount + 1
            End If
        Next i   'NOTE: Starting at 1 and comparing to i + 1 returns bound error
        ReDim stockOutput(1 To outputCount, 1 To 4)   'Reallocating stockOutput to exact size needed
        
        '-----------------------------------------
        'MAIN ALGORITHM FOR OUTPUT CALCULATIONS
        '-----------------------------------------
        Dim startIndex As Long
        startIndex = 1   'Index of 1st instance where stock was found, used for calculations
        
        Dim j As Long
        j = 1   'Incremental counter for position in stockOutput, independent of i (1 --> stockOutputSize)
        
        Dim totalVolume As LongLong
        totalVolume = 0   'Sets initial volume for first loop to 0 for recursive addition
        
        For i = 2 To UBound(stockInput, 1)
            totalVolume = totalVolume + stockInput(i - 1, 7)   'ALWAYS EXECUTES - Sums volumes with same stock name
            
            If (stockInput(i, 1) <> stockInput(i - 1, 1) Or i = UBound(stockInput, 1)) Then
            'Stock change or end is met: Calculate and add values to stockOutput
                
                stockOutput(j, 1) = stockInput(i - 1, 1)   'Logs previous stock ticker
                stockOutput(j, 2) = stockInput(i - 1, 6) - stockInput(startIndex, 3)   'Logs Quarterly Difference
                stockOutput(j, 3) = stockOutput(j, 2) / stockInput(startIndex, 3)   'Logs Quartely % Difference
                stockOutput(j, 4) = totalVolume   'Logs Total Stock Volume
                
                '----------------------------------------------------
                'Creating table: Greatest values of % differences and volumes
                '----------------------------------------------------
                If j = 1 Then   'If no other entries exist then...
                    stockStats(1, 1) = stockOutput(1, 1)   'Storing 1st entry ticker names in all slots
                    stockStats(2, 1) = stockOutput(1, 1)
                    stockStats(3, 1) = stockOutput(1, 1)
                    
                    stockStats(1, 2) = stockOutput(1, 3)   'Storing 1st entries of differences in all slots
                    stockStats(2, 2) = stockOutput(1, 3)
                    stockStats(3, 2) = stockOutput(1, 4)
                Else
                    If stockStats(1, 2) < stockOutput(j, 3) Then   'Compares for greatest % increase
                        stockStats(1, 2) = stockOutput(j, 3)   'Writes over previous values if cond. met
                        stockStats(1, 1) = stockOutput(j, 1)
                    End If
                    If stockStats(2, 2) > stockOutput(j, 3) Then   'Compares for greatest % decrease
                        stockStats(2, 2) = stockOutput(j, 3)   'Writes over previous values if cond. met
                        stockStats(2, 1) = stockOutput(j, 1)
                    End If
                    If stockStats(3, 2) < stockOutput(j, 4) Then   'Compares for greatest volume
                        stockStats(3, 2) = stockOutput(j, 4)   'Writes over previous values if cond. met
                        stockStats(3, 1) = stockOutput(j, 1)
                    End If
                End If
                totalVolume = 0   'Reset volume after every new stock is found
                startIndex = i   'Set new starting stock index for quarterly comparisons
                j = j + 1   'Increment j by 1 for next stockOutput entry
            End If
        Next i
        
        '-------------------------------------
        'PRINTING AND FORMATTING OF NEW DATA
        '-------------------------------------
        'IMPORTANT: About using j as printing index...
        '   j holds +1 value due to nature of algorithm, and printing is desired starting on the 2nd row
        '   Therefore: j is already the adjusted row # for stockOutput
        '***************************************************************************************************
        With ws
            .Range("I1").Value = "Ticker"   'Headers for stockOutput (Columns 9-12)
            .Range("J1").Value = "Quarterly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            .Range("I2:L" & j).Value = stockOutput   'ALL OUTPUT DATA PRINTED FROM THIS ARRAY
            .Range("K2:K" & j).NumberFormat = "0.00%"   'Set % difference column to desired format with %
            
            'Conditional formatting Macro
            '--------------------------
            Dim conditionRange As Range
            Set conditionRange = .Range("J2:J" & j)   'Set variable = range needs formatting

            With conditionRange   'Using "With" for ease of access
                .ClearFormats   'Clears any existing formatting - Zero values are blank
                .FormatConditions.Delete   'Clears any existing conditional formatting too
                
                With .FormatConditions.Add(xlCellValue, xlGreater, "=0.00").Interior   '% > 0 Set to green
                    .ColorIndex = 4
                End With
                With .FormatConditions.Add(xlCellValue, xlLess, "=0.00").Interior   '% < 0 Set to red
                    .ColorIndex = 3
                End With
                .NumberFormat = "#0.00"   'Set quarterly difference to 0.00 decimal placing
            End With
            
            '3x3 Section (Columns 15-17) for stockStats
            'Row Headers
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
            'Column Headers
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"
            .Range("P2:Q4").Value = stockStats   'STOCK SUMMARY PRINTED FROM THIS ARRAY
            .Range("Q2:Q3").NumberFormat = "0.00%"   'Sets "Biggest % Increase/Decrease" formatting
            
            .Columns("A:Q").AutoFit   'After a Worksheet is done being built, autofit all columns
        End With
    Next ws
    
    '##########################################################################
    '# Title: ChatGPT's suggested use of Application.<Property> temporarily turned off Excel
    '#           Worksheet properties that now need to be re-activated
    '# Author: OpenAI
    '# Date: 2024
    '# Code Version: ChatGPT-4o
    '# Availability: https://www.chatgpt.com
    '##########################################################################
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    '(--------------------------)
    '(  CHECKING FOR SPEED HERE   )
    '(--------------------------)
    endTime = Timer
    elapsedTime = endTime - startTime
    MsgBox ("Success! Macro done running!")
    MsgBox ("Execution time: " & elapsedTime & " seconds.")
    '(--------------------------)
    
End Sub





