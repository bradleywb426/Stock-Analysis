# An Analysis of Stock Options

Analysis of a selection of stack data from 2017 and 2018 to investigate investment options

# Overview

Steve's parents are interested in investing in stocks for green energy companies, specifically *DAQO New Energy Corp*. Steve is also looking to diversify his parents investments in other green energy companies. After creating end-user friendly analyses for DAQO stocks as well as the other stocks for the two years of data, Steve is looking to expand the dataset to more stocks over more years. This requires a refactoring of the VBA code to be faster, clearer, and more efficient (as detailed below).

# Results

## Stock Performaces

### 2017 Stock Performance

In 2017, most stocks had increases in yearly returns. Some stocks, like DAQO, had returns of almost +200%, while RUN only has returns of around +6%. Only TERP had a negative return, which was around -7%. Most stocks appear to be good investments based on the 2017 data, particularly DAQO given the large increase on returns. However, trends from the 2018 data show a different story.

<img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Results.PNG>

### 2018 Stock Performance

In 2018, most stocks had negative yearly returns. In fact, only two stocks had postive returns, ENPH and RUN, both of which had around +80% returns. This implies that both stocks may be good investments with consitent returns. Some of the stocks with small negative returns, like VSLR, may be also good investments with the occasional small dip in returns. One of the most notable negative returns is TERP, which has had relatively consistent negative returns for 2017 and 2018. 

<img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Results.PNG.jpg>

## Code Execution Times

### Original Code Times

The orignal code for the Stock analyses performed the analyses relatively quickly. The data for both 2017 and 2018 was analyzed in around .77 seconds. 

<img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/Green_Book_2017.PNG width=375> <img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/Green_Book_2018.PNG>

### Refactored Code Times

The refactored code ran far faster, with each year being both analyzed and formatted in .15625 seconds. The refactored code shaved around .6 seconds off from the original code. 

<img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG> <img src=https://github.com/bradleywb426/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG>

## Refactored Code

There were several key differences between the original code and the refactored code that improved the analysis performance. One of the primary changes was the creation of the ticker index. This index is used to refer to specific tickers in the arrays for tickers, volume, and starting and ending prices, and later used to populate the Results chart with the correct data for respective tickers. This reduces the processing time by combing through each bit of data once, and creating a condition to change the index when appropriate. The creation of arrays for the volume and prices also saves time by not having to update/write the values for each part after checking each cell with each ticker, instead the desired values are simply saved into an array to be pulled from later.

# Summary

Summary: In a summary statement, address the following questions.

-What are the advantages or disadvantages of refactoring code?
  - ADV: simpifies and speeds code
  - ADV: improved readability
  - ADV: fixes missed errors
  - DIS: can be frustrating
  - DIS: takes extra time
  - DIS: not always effective

-How do these pros and cons apply to refactoring the original VBA script?
