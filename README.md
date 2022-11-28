# VBA_Challenge

## Overview
This challenge anazlyzes data in green stocks in order for soeone to make informed decisions as to whether or not to invest in a stock.  I used VBA in order to ouptut total daily volume and return percentages of each stock.  

### Purpose
The purpose of this analysis is to alllow someone ot see the detailed return reports of any specific year.  This will allow an individual to invest in the highest performing stock.


## Results

### Refactored Code

    '1a) Create a ticker Index
    
    Dim tickerIndex As Single
    
        tickerIndex = 0
        
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For j = 0 To 11
     tickerVolumes(j) = 0
    Next j
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 7).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i - 1, 2).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickerIndex Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 7).Value
                
            End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
              If Cells(i + 1, 2).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickerIndex Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 7).Value

            '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                
            End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
            Cells(4 + i, 1).Value = tickers(tickerIndex)
            Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
            Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1

  

## Summary
