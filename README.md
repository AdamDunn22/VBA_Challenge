# VBA_Challenge
## Purpose
We have been hired to analyze the stock information bewtween the years 2017 and 2018, and to try and refractor the code to help reduce the run time it takes for the code to analyze the data.

## Results

### Code
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        'If  Then
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
        'End if
        
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

### Output for 2017 and 2018
Reviewing the results we can see that all but one of the stocks available is in the green for 2017, while for 2018 all but two of the stocks are negative. So we can see that 2017 was a better year compared to 2018.

## Summary
### Advantage and Disadvantages of Refractored Code
Advantages
* Remove excess code to help make it easier to run and understand
* Fix bugs, and organize code
* Faster to write the code

Disadvantages
* Can break the code if it was not structured properly
* May create bugs which can be risky
* The process requires skill, time, and knowledge of the code and VBA code

### Advantages and Disadvantages of Original and Refractored Code
* Advantages
  * The refratored code is easier to read and runs a lot quicker over the original code. Refractoring is a process in optimizing the exisiting code, it allows the code to run faster, easier to read.
* Disadvantages
  * The script only allows us to search one year at a time. One of the disadvatages of refaractoring is while refractoring we may miss codes and break the codes; in this situation the process can be risky and time consuming.
