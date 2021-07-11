# stocks-analysis


## Overview of Project 
In this Project, I am helping Steve analyze stock data and create a report showing how the stock "DAQO" has been doing over the last few years. He needs this data to share with his parents as they want to invest in stock through him.

### Purpose
The purpose of this project is to analyze all the data from all the stock in 2017 and 2018, refactor the data to show how well the "DAQ0" stock has been performing

## Background
Steve just graduated from college with a finance degree and his parents want to be his first clients. They are interested in renewable energy so they want to invest in the stock "DAQ0". In order for Steve to convince his parents he needs a report that shows how "DQ" has performed in comparison to other stock over the last few years.


## Results
Based on my results, I was able to get a table that shows the average return of  stock for 2 years (2017) and (2018). The table is supposed to match with the tables from the module. However, mine was quite different. All of my stock were in the red (negative values). My values were also slightly different from that in the module. I made some errors and unfortunately able to figure out what I did wrong. Following the questions, I was able to loop through the tickers to get the average return
example: For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
	Cells(4 + i, 1).Value = tickers(i)
	Cells(4 + i, 2).Value = tickerVolumes(i)
	Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
This in addition with other coding scripts produced a table with 3 headers similar to the module. In the end I was able to get the code performance for the stocks in [2017](https://github.com/somtoesomeju/stocks-analysis/blob/main/VBA_challenge_2017.png) and [2018](https://github.com/somtoesomeju/stocks-analysis/blob/main/VBA_challenge_2018.png## Summary)
To summarize, I was unsuccessful in replicating the code as it was from the module, although I was able to produce a table with the values. Refactoring code is a great way to get a more detailed analysis of what is being looked for in the excel sheet. It also significantly reduces the time it would take to go through the cells and calculate these values manually. However, if not done correctly it can lead to many errors. This was my experience in this assignment.

In regards to this assignment, the main advantage of refactoring the code is that it shows there are multiple ways to reach the end result. From the module, a different set of code was produced to give the refactored table. Using the vba template, we were tasked in refactoring it differently. This produced similar results. The disadvantage though, like in my case is that if not done properly it can further complicate the data.

