# stock-analysis VBA
## Overview of Project
The purpose of this project is to analyse the stock data of 12 different stock options, and determine which stocks are worth investing in, 
and what the outcomes for years 2017 and 2018.

## Data
The data provided is two different charts (one for each year) to see the returns and total daily values for both years. 

## Results
As we will be able to see in the file, the year of 2017 did significantly better in the returns section, than the 2018. Before I refactored the code I went over the instructions of what was needed to make this code as easy to read as possible. There was already an structure for me to start with and below is how my code looks like in the file.  

    
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
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
    
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3d Increase the tickerIndex
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerIndex = tickerIndex + 1
            
        End If
        
    Next i

     '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
     For i = 0 To 11
    
      Worksheets("All Stocks Analysis").Activate
    
     Cells(4 + i, 1).Value = tickers(i)
     Cells(4 + i, 2).Value = tickerVolumes(i)
     Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
     Next i
        
        
 # Summary
 ## Advantage and Disadvantages of Refactoring Code
 One of the main advantages and main purpose of refactoring code, is being able to improve the internal code without afecting the external behavior. The simple way to do this is just making small changes at a time and make sure everything is running smoothly, and working. 
 Now, a disadvantage on the other hand, is just how much time it takes to get it done. Sometimes you run into some issues that are just time consuming, and can leave you feeling lost/confused for a while until you finally figure out where everything went wrong.
 
 ## Advantages and Disadvantages of the Original and Refactored VBA Script
 
The main advantage of refactoring an original scrip is just to compare the improvements made to the original code. Although you have to be continuosly testing it, it just makes the original scrip work better and faster, and it's just a good maintanance habit to keep thing running smoothly. 
A disadvantage on the other hand, is then again just time consuming at times, and just the fact that you have to continously run your code to make sure it works properly. 
 

