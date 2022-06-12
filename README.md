# An Analysis of Stock Options

Analysis of a selection of stack data from 2017 and 2018 to investigate investment options of green energy companies.

# Overview

Steve's parents are interested in investing in stocks for green energy companies, specifically *DAQO New Energy Corp* or DQ. Steve is also looking to diversify his parents' investments in other green energy companies. After creating end-user-friendly analyses for stocks of DQ and other green energy companies with two years of data, Steve is looking to expand the dataset to include more stocks over more years. This requires a refactoring of the VBA code to be faster, clearer, and more efficient when it analyzes far more data.

# Results

## Stock Performances

### 2017 Stock Performance

In 2017, most stocks had increases in yearly returns. Some stocks, like DAQO, had returns of almost +200%, while RUN only has returns of around +6%. Only TERP had a negative return, which was around -7%. Most stocks appear to be good investments based on the 2017 data, particularly DAQO given the large increase in returns. However, trends from the 2018 data show a different story.

<img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Results.PNG>

### 2018 Stock Performance

In 2018, most stocks had negative yearly returns. In fact, only two stocks had positive returns, ENPH and RUN, both of which had around +80% returns. Given the consecutive years of positive returns, this implies that both stocks may be good investments that provide consistent returns. Some of the stocks with small negative returns, like VSLR, maybe also be good investments in the long run that had a small decrease in returns, but more data would be required to properly make such assumptions. One of the most notable negative returns is TERP, which has had relatively consistent negative returns for 2017 and 2018 and is likely a company on the decline. 

<img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Results.PNG.jpg>

## Code Execution Times

### Original Code Times

The original code for the Stock analyses performed the analyses relatively quickly. The data for both 2017 and 2018 was analyzed in around .77 seconds, which is sufficiently fast for this data set. However, the original code would have noticeably lagged when applied to potential larger data sets. 

<img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/Green_Book_2017.PNG width=375> <img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/Green_Book_2018.PNG>

### Refactored Code Times

The refactored code ran far faster than the original, with each year being both analyzed and formatted in .15625 seconds. The refactored code shaved around .6 seconds off from the original code, a major improvement from the original, and would likely perform far better than the original when applied to larger data sets. 

<img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG> <img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG>

## Refactored Code

### Ticker Index

There were several key differences between the original code and the refactored code that improved the analysis performance. One of the primary changes was the creation of the ticker index (See code 1A) to refer to specific tickers in the arrays for tickers, ticker volumes, and starting and ending prices. This index is akin to the tickers but is coded to automatically increase in value with each new ticker (See code 1B) so as to not mix data. The use of the tickerIndex also negates the use of a nested loop, which slowed the original code by combing the data several times as it checked each cell for each ticker. Note that the code in the following section and in future sections are taken out of order to show specific lines:
```
1A)
Dim tickerIndex As Single
    tickerIndex = 0
    
...
    
1B)
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    tickerIndex = tickerIndex + 1
End If
```
Code 1A creates the variable "tickerIndex" and initializes it to zero. Code 1B then checks if the value in the current cell of the first column, Cells(i,1).Value, equals the value of the next cell in the first column, Cells(i+1,1).Value, and the i is from the for loop this condition is nested in, in which i = 0 to 11. If the value of the cells is equal, then the code moves onto the next cell and checks all current conditions for that cell. However, if the values do not match, then the tickerIndex is increased by one and the conditions are checked for that next cell.

### Other Arrays

This index was used to store data in three arrays (See code 2A): ticker volumes, ticker starting prices, and ticker ending prices. Storing data in these arrays using conditionals (See code 2B) allows the code to be parsed quicker as it does not have to check to print and then print data for each cell after checking each ticker. Instead, data can be pulled from the array at a later time to populate a table (see code 2C).
```
2A)
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

2B)
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
End If

If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
End If

2C)
For i = 0 To 11   
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = (tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex)) - 1
Next i
```
Code 2A creates arrays of 12 empty values that are populated later. Code 2B has three parts, the first of which adds the volume of the current row of cells to the total volumes for the current ticker based on the tickerIndex value. Then the code checks if the previous cell is the first cell of a ticker by checking if the ticker value (based on the tickerIndex) of the previous cell, Cells(i-1, 1).Value, equals the ticker value of the current cell, Cells(i, 1).Value. If the current cell is the first cell of a ticker, then the starting price is recorded into the tickerStartingPrices array for the tickerIndex Value. If the cell is not the first cell for the ticker, then this condition is skipped. The same general process is followed for the tickerEndingPrices array, where the code checks to see if the current cell is the last cell for a ticker, Cells(i+1, 1)>Value, and if it is, copy that value into the array. Lastly, Code 2C does a quick for loop to add the values stored in the aforementioned arrays to a table created in one of the sheets.

To view the code and sheet in context and as a whole, please go to the Excel file: [VBA Challenge](https://github.com/bradleywb426/stock-analysis/blob/main/VBA_Challenge.xlsm).

# Summary

### Pros and Cons of Refactoring

Refactoring codes can be a very useful skill to possess, but refactoring itself has many pros and cons to it. Refactoring code often simplifies it, allowing for code that can run faster and be more easily understood (especially when the code has comments in it). Refactored code can often also fix unseen errors or complications in the code that were missed before. On the other hand, refactoring takes extra time to make work and depending on the original code and the project can require a rather large amount of time. Refactoring can also be frustrating, largely in part of the extra time needed but also in the search for how to simplify the code. Refactoring code is not always a guaranteed large improvement as some code can not be improved much further despite spending a lot of time and effort on it.  

### How This Applies to the Refactored Code Here

The pros and cons discussed above can be applied to the refactoring of the original VBA script for this project. After refactoring, the code became easier to read and ran faster than the original code while also formatting the table created with the code. The process of refactoring took extra time to accomplish, however, and troubleshooting the refactoring process also proved mildly frustrating at times.
