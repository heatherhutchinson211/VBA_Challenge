# VBA_Challenge

## Overview
This challenge anazlyzes data in green stocks in order for soeone to make informed decisions as to whether or not to invest in a stock.  I used VBA in order to ouptut total daily volume and return percentages of each stock.  

### Purpose
The purpose of this analysis is to alllow someone ot see the detailed return reports of any specific year.  This will allow an individual to invest in the highest performing stock.


## Results
I intially created a code to run this analysis that was able to execute code of rthe 12 stocks provided;  however, I then refactored the code in order ot perform bettter over a multitude ofstocks in a single loop and a much quicker execution time.  

I did this by creating a nested for loop that runs through each stock and outputs the data for that stock individually. 

However, I did run into a problem with my nested for loop that wasn't allowing the loop to move onto the next index.  I tried my best to debug and figure out where the problem lied, but unfortunately I wasn't ablw to find the mistake.  With that being said, I was still able run the code, and the refactored code ran in significantly less time that the original code.

Original code run time:
<img width="257" alt="Screenshot 2022-11-27 at 9 55 46 PM" src="https://user-images.githubusercontent.com/117620028/204216084-b56fee0b-79dd-4df4-874c-f8b0a3bf20f1.png">

Refactored code run time:
<img width="255" alt="Screenshot 2022-11-27 at 9 57 01 PM" src="https://user-images.githubusercontent.com/117620028/204216097-7065382c-ddd9-475f-80e0-a8e779e90dbe.png">


### Refactored Code
This is the code that I used after refactoring. 

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
Overall, it is clear that refactoring code has great benefits.  In this situation, the code was much shorter and quicker to write, as well as allowing much larger data sets to be incoroporated.  Not only would someone be able to run this code over 12 stocks, but it would also be easy to use this code for an even larger analysis.  

In general, refactoring code can be very beneficial when working with large data sets.  It can can make your code more efficient and quicker to run.  However, it can sometimes bring disadvantages if your current code is already running smoothly.  
