#README.md STILL IN PROGRESS

# VBA-challenge
# UTA-VIRT-DATA-PT-11-2024-U-LOLC

## Listed Sources

## MODULE PURPOSE:
      Module takes the input of an Excel workbook with (potentially) multiple worksheets within it and creates two
      new data tables in each one complete with new formatting and headers. Workbook should only update after
      execution of module. Subroutine is built using arrays to handle large quantities of data!

## DETAILED DESCRIPTION OF MODULE:
      Built to work with daily entries of various stocks and their associated values, listed in one or more worksheets.
      Suboutine takes input of a workbook and copies organized data (with its dimensions) to the temporary memory
      of local computer. Subroutine works in one worksheet at a time until output for each one is complete.
      Looping within inputted data, it only pauses at new stock tickers to compute statistics of prior range of dates.
      Stats are stored in two new 2D arrays on computer by grouping of stock ticker. After new arrays are built, data
      is printed back to each worksheet with conditional formatting. Permits UI to resume updates before subroutine end.

## CURRENT LIMITATIONS:
      Runs only with Microsoft Excel Workbooks.
      Stock tickers don't have to be in alphabetical order, but they MUST be grouped by ticker name.
      Stock tickers in the same group MUST be in exact chronological order (Earliest date --> Latest Date).
      Input format must be: {Rows} x {Columns} --> {Ticker, date, open, high, low close, volume} x {Daily stock entries}
      Input data is NOT verified/formatted/cleansed by module - adding said functionality has not been addressed.

## CURRENT METHOD PATH:
      <Run Macro>
      --> Disable application update features --> Begin loop through each worksheet
      --> Copies data and perserves formatting/arrangement in same 2D array dimensions
      --> Runs data through tests in local memory --> Stores results in two new 2D arrays
      --> Two new arrays with headings and formatting printed back to workshees
      --> Loop to next worksheet
      --> Success! Re-enable application update features to see changes
      <End Macro>

