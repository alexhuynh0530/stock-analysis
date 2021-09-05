# An Analysis of Green Energy Stocks using Excel and VBA

## Overview of Project

### Purpose

This analysis is to help Steve, a recent graduate with a finance degree. His parents are passionate about green energy and have decided to invest all their money into DAQO New Energy Corp ($DQ) without doing much research. Steve has promised to look into DAQO stock for his parents but is concerned about diversifying their funds, so he wants to analyze a handful of green energy stocks in addition to DAQO's stock.

Steve has created an excel file containing the stock data and has asked us to help him analyze it. By using Excel and VBA, we automate the analysis using code and built-in macros to run scripts that finds the total daily volume and yearly return for each stock. The results will help Steve show his parents the performance of the green energy stocks and help them decide if DAQO is a good investment. 

In this anlysis, we will use the results of our script to compare the stock performance between 2017 and 2018. In additon, we have refactored the original code and will analyze and compare the new refactored code versus the old code. We will also discuss the difference in execution times between the two.

## Results

### Comparing the stock performance between 2017 and 2018

#### 2017

![VBA_Challenge_2017_results.png](https://github.com/alexhuynh0530/stock-analysis/blob/main/Resources/VBA_Challenge_2017_results.png)

#### 2018

![VBA_Challenge_2018_results.png](https://github.com/alexhuynh0530/stock-analysis/blob/main/Resources/VBA_Challenge_2018_results.png)

When comparing the performance of the green energy stocks shown above, we can see that 2017 had better returns across the board than 2018. You'll see that in 2017, DAQO outperformed the group of green energy stocks, returning 199.4% on the year. TERP was the only stock that had negative returns for 2017, although it was a loss of less than 10%.

In 2018, only 2 stocks had a green year, ENPH and RUN, both returning more than 81% while the rest of the group, including DAQO New Energy Corp ($DQ), had negative returns. In addition, DAQO had the worst returns out of the whole list of green energy stocks in 2018 returning -62.6%. 

To improve the analysis, please see the average return from 2017 and 2018 below.

#### Average Return from 2017 and 2018

![VBA_Challenge_avg_return.png](https://github.com/alexhuynh0530/stock-analysis/blob/main/Resources/VBA_Challenge_avg_return.png)

As you can see, the top 3 stocks that had the best returns was ENPH, SEDG, and DQ returning 105.7%, 88.4%, and 68.4% respectively. You can conclude that ENPH is the best stock to own from the data given. In addtion to having the best average returns from 2017 and 2018, ENPH also had consistent gains in both years returning 129.5% and 81.9%. 

When looking at DQ stock, although it made the top 3 stocks with highest returns, you'll see that it is highly volatile. It had the highest return in 2017 but the worst returns in 2018. This doesn't seem like an investment that Steve's parents should invest in especially at their age.

Please note, there are some limitations to this dataset as the data is only limited to 2017 and 2018. Global news, company specific news, and many other factors could have affected the performance of a stock in a particular year. Therefore, analyzing data with more years would improve the anlaysis.

### Comparing the execution times of the original script and the refactored script

#### Origianl Script Run Times

![VBA_Challenge_2017_oldcode.png](https://github.com/alexhuynh0530/stock-analysis/blob/main/Resources/VBA_Challenge_2017_oldcode.png)

![VBA_Challenge_2018_oldcode.png](https://github.com/alexhuynh0530/stock-analysis/blob/main/Resources/VBA_Challenge_2018_oldcode.png)

#### Refactored Script Run Times

![VBA_Challenge_2017.png](https://github.com/alexhuynh0530/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018.png](https://github.com/alexhuynh0530/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

From the screeshots above, you'll see that the refactored script runs much faster versus the original script. 

### Overview of original script and the refactored script

#### Original Script

```
'3) Prepare for the analysis of tickers.
'3a) Initialize variables for the starting price and ending price.
    
    Dim startingPrice As Double
    Dim endingPrice As Double

'3b) Activate the data worksheet.

    Sheets(yearValue).Activate

'3c) Find the number of rows to loop over.

    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through the tickers.

    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0

'5) Loop through rows in the data.

        Sheets(yearValue).Activate
        For j = 2 To RowCount

'5a) Find the total volume for the current ticker.

            If Cells(j, 1).Value = ticker Then
                'increase totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(j, 8).Value
            End If

'5b) Find the starting price for the current ticker.

            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                'set starting price
                startingPrice = Cells(j, 6).Value
            End If


'5c) Find the ending price for the current ticker.

            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                'set ending price
                endingPrice = Cells(j, 6).Value
            End If

        Next j
        
'6) Output the data for the current ticker.

    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i
```

In the original script, we loop through the tickers and use a nested loop to loop through the rows of data. To calculate total volume for each ticker, we look at the 1st column to see if it matches a specific ticker and if it does then we add the volume in the 8th column for that row. We continue adding the volume for each row that ticker is a match. We also store the starting price data as a "Double" data type and look at the 1st column to see if it matches the specified ticker for that loop iteration while making sure the ticker before that row does not match. We also store the ending price as a "Double" data type and match the ticker while making sure the ticker after that row does not match. Before we loop through the next ticker iteration, we print out the data for ticker, total volume, and calculate the return by dividing the ending price by the starting price.

#### Refactored Script 

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
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                'set starting price
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
            'End If
            End If

            '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                'set ending price
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                
            'End If
            End If
        
        Next i

    'Next tickerIndex

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
In the refactored script, we create three output arrays to store total volume, starting price, and ending price. We use the tickerIndex to access the correct index across the ticker array and three different arrays. We create a for loop to initialize tickerVolumes to zero. We then create another loop to loop over all the rows while increasing the volume for the current ticker which is stored in the tickerVolumes array. We look for the starting price and ending price in the same way as the original code, however, we store the data in the tickerStartingPrices and tickerEndingPrices arrays. If the ending price is found then we increase the tickerIndex by 1. After all the data is stored in the arrays, we output the data using a for loop and referencing the arrays to print the ticker, total volume, and calculate the return by dividing the ending price by the starting price.

#### Conclusion

You can conclude that with the use of the arrays in the refactored script, it has improved the run time versus the original script.

### Summary

#### Advantages of refactoring code in general

- More efficient
- Faster run times
- Less memory
- Improving logic and makes it easier for future users to read

#### Disadvantages of refactoring code in general

- Introducing new bugs

#### Advantages of the original VBA script

- The code works even if the tickers were not in alphabetical order

#### Disadvantages of the original VBA script

- Slower run time

#### Advantages of the refactored VBA script

- Faster run time

#### Disadvantages of the refactored VBA script

- If tickers were not in alphabetical order the code does not work
