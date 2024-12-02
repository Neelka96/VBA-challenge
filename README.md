> [!Warning]
> README.md STILL IN PROGRESS

# README: VBA-challenge
> Module Path: VBA-challenge/2oStockRoutine.bas
> GitHub Path: VBA-challenge/README.md
> 
> **Last Updated: December 1st, 2024**

# Important Information
Script is a macro written in VBA (Visual Basic for Applications) for use solely with Microsoft Excel files types and has only been tested with .xlsm and .xlsx file formats.

**Made for EdX & UT Data Analytics and Visualization Bootcamp: Cohort UTA-VIRT-DATA-PT-11-2024-U-LOLC.** This is only the second module completed in the course!

README.md was written using tips and tricks from [GitHub Docs](https://docs.github.com/en/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax).

For citation information please see [VBA-challenge Citations](#vba-challenge-citations)!

## CURRENT METHOD PATH:
```
<Run Macro>
--> Disable application update features
--> Begin loop through each worksheet
--> Copies data and perserves formatting/arrangement in same 2D array dimensions
--> Runs data through tests in local memory
--> Stores results in two new 2D arrays
--> Two new arrays with headings and formatting printed back to workshees
--> Loop to next worksheet
--> Success! Re-enable application update features to see changes
<End Macro>
```

## Module Purpose
Module's purpose is to analyze a large amount of data on daily stock values and provide a summary for each one. After an examination of the data, the module creates new tables of data with corresponding formatting and headers. The tables will be filled in with computations that have been run on each grouping of stock tickers to create the summaries. The workbook and any worksheets within it should only update after the module finishes its execution.

## Detailed Description of Module
Built to work with daily entries of various stocks and their associated values listed in one or more worksheets. Suboutine takes the input of a workbook and copies pre-organized data within the workbook (and its dimensions) to the temporary memory of local computer in the form of a variant array. The array data include the stock name, date of entry, the opening, high, low, and closing prices, and volume for each entry. The sample data given by the class has the workbook seperated into quarters and the stocks grouped by tickers which makes comparing groups of stocks easier. 

Looping within the new array, it only pauses at new stock tickers to compute the statistics of the prior range of stock tickers. The module will summarize the data for each stock ticker group and working one quarter at a time until the output is completed. The computations are stored in two new 2D arrays on computer by grouping of stock ticker. After new arrays are built, data is printed back to each worksheet with macro-based conditional formatting. Finally, the module permits the Microsoft Excel UI to resume updates before subroutine end.

> [!NOTE]
> [Current subroutine](/2oStockRoutine.bas) is now built using arrays to handle large quantities of data with rapid efficiency! [Previous subroutine](/1oStockRoutine.bas) ran by continuously accessing Microsoft Excel Ranges in the workbook and storing all temporary data within the sheet, however it computed at only 20% of the speed of current mode.

## Current Limitations
- [ ] Runs only with Microsoft Excel Workbooks.
- [ ] Stock tickers don't have to be in alphabetical order, but they MUST be grouped by ticker name.
- [ ] Stock tickers in the same group MUST be in exact chronological order (Earliest date --> Latest Date).
- [ ] Input format must be: {Rows} x {Columns} --> {Ticker, date, open, high, low close, volume} x {Daily stock entries}
- [ ] Input data is NOT verified/formatted/cleansed by module - adding said functionality has not been addressed.

# VBA-challenge Citations
For writing [2oStockRoutine.bas](/2oStockRoutine.bas) source code and its supporting README.md documentation multiple sources were used acrossed the web including [ChatGPT](#chatgpt-suggested-optimization), [Microsoft Learn](#microsoft-learn-vba-reference-guide), and [GitHub Docs](#readme-md-citations).

## Source Code Citations
1. ChatGPT Optimization Suggestion Using Application Object Properties [^1]
   - Desciption: ChatGPT's algorithm suggested use of setting Application.Property = Value temporarily turns off Excel Worksheet properties that run during macro use until subroutine is done modifying data to improve latency.
   - Author: OpenAI
   - Date: 2024
   - Code Version: ChatGPT-4o
   - Availability: https://www.chatgpt.com
[^1]: [ChatGPT-4o by OpenAI (2024)](https://www.chatgpt.com) used for optimization help:
  Suggested modifying module [1oStockRoutine.bas](/1oStockRoutine.bas) setting Application.<Property> values
```
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

<Subroutine Core Goes Here>

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
```
2. ChatGPT Optimization Suggestion Using Variant Arrays [^2]
   - Description: ChatGPT suggested use of RAM (variant type array method) for storing, accessing, computing, comparing, and printing data. Implemented for its speed.
   - Author: OpenAI
   - Date: 2024
   - Code Version: ChatGPT-4o
   - Availability: https://www.chatgpt.com
[^2]: [ChatGPT-4o by OpenAI (2024)](https://www.chatgpt.com) used for optimization help:
  Suggested modifying module [1oStockRoutine.bas](/1oStockRoutine.bas) to utilize arrays
```
Dim stockInput As Variant   'Array will hold all original data (aka Input)
'Has dimensions of: {# Columns} x {# Rows} or more precisely...
'{Ticker, date, open, high, low close, volume} x {Distinct stock entry dates}
    
Dim stockOutput() As Variant   'Array will hold all output values - ReDim'd later at exact sizing
'Has dimensions of: {# Columns} x {# Rows} or more precisely...
'{Ticker, quarterly difference, quarterly percent difference, total volume} x {Distinct stock tickers}
    
Dim stockStats(1 To 3, 1 To 2) As Variant   'Array will hold largest values within stockOutput() for output also
'Has dimensions of: {# Columns} x {# Rows} or more precisely...
'{Ticker, value} x {Greatest % increase, greatest % decrease, greatest volume}
```
3. Microsoft Learn VBA Reference Guide [^3]
   - Description: Information on UBound(), With, ReDim declaration, variant data type, and conditional formatting
   - Author: Microsoft Learn/@o365devx/@AlexJerabek/@kbrandl/@OfficeGSX/@Saisang
   - Date: 2021-2022
   - Availability: https://learn.microsoft.com/en-us/office/vba/api/overview/
[^3]: [Microsoft Learn VBA Reference Guide](https://learn.microsoft.com/en-us/office/vba/api/overview/) used for help with understanding built-in formula syntax:
   Formulas include UBound(), With var As Type, ReDim var As Type, and Dim var As Variant, and using Range.FormatConditions.Add()
```
For i = 2 To UBound(stockInput, 1)   'Looping from 2nd index to last index of stockInput()
      If stockInput(i, 1) <> stockInput(i - 1, 1) Then   'Checks if prior value is different
            outputCount = outputCount + 1
      End If
Next i   'NOTE: Starting at 1 and comparing to i + 1 returns bound error
ReDim stockOutput(1 To outputCount, 1 To 4)   'Reallocating stockOutput to exact size needed

...

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
```

## README.md Citations
1. https://lib.pstcc.edu/csplagiarism/citation
2. https://integrity.mit.edu/handbook/writing-code
3. https://docs.github.com/en/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax

## List of Subroutine SoftLinks
[Link to Subroutine 2o](/2oStockRoutine.bas)
[Link to Subroutine 2o With Timer](/2oWithTimer.bas)
[Link to Subroutine 1o](/1oStockRoutine.bas)
