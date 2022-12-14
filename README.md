# Stock-analysis With VBA

## Overview of Project

Our client, Steve, is helping his parents find investments. He would like to analyze stock market data over the last several years to find which green stocks show the best growth. We've created a workbook using VBA to analyze the change in certain stocks over a year's time and provide us with the total yearly volume and yearly return for each stock.  Now, we have refactored the code in order to collect the same information more efficiently. With a faster run time, Steve can explore even larger datasets at a faster pace and help his parents choose their investments wisely.

### Dataset

We were presented with data for several stocks over the years 2017 and 2018.  Each dataset by year includes the ticker IDs for individual stocks, the date, opening and closing prices for each date, the daily high and low prices, the adjusted closing price, and the volume of the stock.  Steve identified 12 stocks he wants to research. In order to get the percentage growth of each stock, we need to find the starting price for the year and the ending price for the year and calculate the percentage growth. 

## Results

Using VBA, we wrote a script to index the 12 stock tickers, loop through the dataset, and pull the starting prices, ending prices, and total volumes for each ticker based on the year input for the analysis. Each ticker ID, their total volumes, and yearly returns were then pulled onto a new sheet and formatted to make it easier to identify the stocks with the highest return. Running the two different years through the script, we find that 2017 was a good year for almost all the stocks while only 2 stocks, ENPH and RUN, showed positive returns in 2018.  From this, we may suggest Steve's parents invest in ENPH and RUN because of their continued growth across both years.

![VBA_Challenge_2017.png](https://github.com/ChallahBack83/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![VBA_Challeng_2018.phg](https://github.com/ChallahBack83/Stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)


### Refactored Code

In order to make this code more effective and usable across larger datasets and over several years, Steve has asked us to refactor the code. For this challenge, we took the existing code and followed the instructions to work through the refactoring by creating arrays, variables, and for loops.  Below is the refactored code used within the instructions provided:

```
  '1a) Create a ticker Index
       
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    
    For j = 2 To RowCount  
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
                
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
        End If
         
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
               
        If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                   
           '3d Increase the tickerIndex.
            
           tickerIndex = tickerIndex + 1
        End If 
        
    Next 
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```

### Run Time

After changing the code, I reran the original script and compared the run time for both years between the original and refactored code.  It is very obvious when you look at them side by side that the new script greatly reduced both run times.

#### Original Run Time for 2017

![Original_Run_Time_2017](https://github.com/ChallahBack83/Stock-analysis/blob/main/Resources/VBA_Run_Time_1_2017.png)

#### New Run Time for 2017

![New_Run_Time_2017](https://github.com/ChallahBack83/Stock-analysis/blob/main/Resources/VBA_Run_Time_2017.png)

#### Original Run Time for 2018

![Original_Run_Time_2018](https://github.com/ChallahBack83/Stock-analysis/blob/main/Resources/VBA_Run_Time_1_2018.png)

#### New Run Time for 2018

![New_Run_Time_2018](https://github.com/ChallahBack83/Stock-analysis/blob/main/Resources/VBA_Run_Time_2018.png)

## Summary

### Advantages and Disadvantages of Refactoring Code

The advantages of refactoring code are that it increases functionality, making the code usually more efficient and run faster. It also makes the code much more readable for future coders. Someone else will be able to come in and understand what each step more clearly. This helps them then to apply their own knowledge to build upon the code and hopefully improve it even further. However, refactoring code is not without it's disadvantages. For instance, you may introduce new bugs or errors to previously cleanly running scripts. It also may not actually make a long script shorter depending on the steps you need to take to make the refactoring functional.

### Advantage and Disadvantages of Refactoring for Stock-Analysis Challenge

As we can see for the results of this challenge, one of the biggest advantages of refactoring code is making the code run more efficiently. This means the script runs faster and increases the amount of data that can be analyzed in a certain set of time. For both 2017 and 2018, the run time was approximately .50 seconds faster.  The only disadvantage was the time it took to recreate new arrays and variables in order to create that efficiency.

