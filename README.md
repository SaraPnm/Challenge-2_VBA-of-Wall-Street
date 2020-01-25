# Stock Analysis of Wall Street
## Description:
Module 2 challenge is in [VBA of Wall Street_green_stocks](https://github.com/SaraPnm/Challenge-2_VBA-of-Wall-Street/raw/master/VBA%20%20of%20Wall%20Street_green_stocks.xlsm)
In this module, we're helping Steve to analyze the stock data for his parents' investment.


### Challenge 2
Sheet *"2017"* and sheet *"2018"* contain the stock data, and sheet *"All Stock Analysis"* is the output sheet including two buttons: *"Clear Worksheet"* to clean the entire sheet, and *"Run Analysis for All Stocks"* to run the analysis. 

In the class excericeses, we went over the general stock analysis using VBA. The challenge requires us to refactore the VBA code to perform the same analysis in a more readable and efficient way. The actions I took are as follows:

- created dynamic arrays for all the tickers, volumes, starting prices, and ending prices. The reason why I decided to took a step forward to make the arrays dynamic is to give the flexibility to Steve to add more tickers in future analysis.

- set an index to loop through the available tickers taking advantage of the fact that the tickers are all in alphabetical order.

- used the tickerIndex to access the correct index across the four different arrays.

- output all of the information I’ve collected.




