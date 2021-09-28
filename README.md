# Stock Analysis Project
## Overview
After having written a VBA code to perform analyses for Steve of multiple stocks’ performances over two years, Steve was keen to expand the analyses to include entire stock markets over multiple years. Due to the large increase in the number of stocks to analyze, the existing VBA code could take too long to run.

The purpose of this project is therefore to re-look at the existing VBA code and refactor it so that it can churn out the analyses over large number of stocks more efficiently.

We will also review the stock performance of the existing dataset for 2017 and 2018 to evaluate how they have done during those two years.

## Code Refactoring
Original VBA Code 

![Original VBA Code](https://github.com/hwaijiinlee/stock-analysis/blob/main/Resources/AllStocksAnalysis_Original.png)

Refactored VBA Code

![Refactored VBA Code](https://github.com/hwaijiinlee/stock-analysis/blob/main/Resources/AllStocksAnalysis_Refactored.png)

In the original VBA code, we had used nested For loops to cycle through all rows of data to obtain volumes, starting and ending prices of each ticker. This increased the complexity of the code and therefore required more resources to run.

In the refactored code, we created a ticker index variable instead. Using this variable, we avoid using nested For loops to cycle through all tickers by increasing the index after the current ticker’s ending prices were being extracted.

The runtime of the refactored code significantly reduced from about 1.6 seconds (for 2017’s data) and 1.5 seconds (for 2018’s data) to about 0.3 seconds and 0.2 seconds respectively. This reduction bodes well for scaling our analyses to much larger datasets.

Refactored VBA Code Runtime

![Refactored VBA Code Runtime 2017](https://github.com/hwaijiinlee/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![Refactored VBA Code Runtime 2018](https://github.com/hwaijiinlee/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Analyses of Stock Performance
![2017 Stock Performance](https://github.com/hwaijiinlee/stock-analysis/blob/main/Resources/AllStocksAnalysis_2017.png)
![2018 Stock Performance](https://github.com/hwaijiinlee/stock-analysis/blob/main/Resources/AllStocksAnalysis_2018.png)

Comparing the stocks’ performances over 2017 and 2018, it is clear that 2018 was not a good year to sell, especially if one had purchased those stocks back in 2017. Macro-economic factors were the underlying cause of the bearish market -  the US’ trade war with China, rising interest rates, and slowdown in economic growth globally were some notable ones. The stocks’ performances in 2017, however, is reason to be optimistic that should those macro-economic factors be resolved or improved, the likelihood of the stock market returning to a more bullish state is high.

## Summary
The difference between the codes’ runtime was astonishing. The refactored code only took a mere tenth of the runtime of the original code. It is therefore important for experienced or knowledgeable programmers to review codes written by those new to the field to ensure optimal efficiency can be obtained with less complexity. As programmers grow into their roles, it is also crucial that they review codes they have written in the past to see if improvements can be made to them.

Given the advantages of refactoring codes, we have to also bear in mind that codes can be broken probably easier than they can be improved. Unless the person looking to refactor the codes is confident and knowledgeable, codes that are working as intended are better left alone.

In this project, it is obvious that the advantages of refactoring far outweigh the disadvantages given that the original codes are written by people new to the field of VBA programming.


