# Green Energy Stock Analysis 

## Overview of Project

### Purpose
The purpose of this analysis is to assist our client, Steve, by analyzing stock data for different green energy stocks and evaluating their performance in the stock market from 2017 to 2018. Steve will be using our analysis of the total daily volumes and yearly returns for each stock to advise his own clients on which stock they should invest their money in. After completing our initial analysis, we felt it necessary to refactor our initial code to improve its efficiency. This report will summarize our findings and explain whether refactoring improved our code's efficiency in analyzing large quantities of stock data. 

## Results

### Parameters of a Stock's Performance
This analysis evaluated the performance of twelve different green energy stocks by calculating their total daily volumes and yearly returns in 2017 and 2018. 

The total daily volume is a measure of a stock's trade activity, precisely the number of shares traded by stock owners per day. The greater the total daily volume over the year, the greater the stock's movement in the market, which is reflective of how many people are interested and invested in that particular stock. 

The yearly return is the percentage difference in closing price from the beginning of the year to the end of the year. The greater the value, the greater the price increase in the stock and the more money earned by shareholders. 

Please see below for the original VBS script that calculates the total daily volumes and yearly returns for each stock, based on the user's input for which year the user wants to analyze.

Please see below for the refactored VBS script that calculates the total daily volumes and yearly returns for each stock, based on the user's input for which year the user wants to analyze.

Note that the formula structure to calculate the total daily volume and identify the starting/ending prices for the yearly return calculation remains the same between the original and refactored script. Please note the tickerIndex included in the refactored script allowing all three output arrays to access the stock tickers for analysis without looping through the entire spreadsheet for each stock ticker each time and wasting time. 

### Analysis of Stocks' Performance
In 2017, all stocks but "DQ" and "HASI" exceeded 100,000,000 total shares traded daily over the year. All but "TERP" had a net positive yearly return, with "DQ" experiencing a price increase of almost 200%. Overall, 2017 was a good year for most green energy stocks. If we're looking into only this year's analysis to select stocks to invest in potentially, the results indicate that "ENPH" and "FSLR" are great choices because they had two of the highest total daily volumes and experienced yearly returns of over 100%. 

Please refer to the below-attached data table presenting our analysis of the stocks' performance in 2017. 

In contrast to 2017, many stocks' performances declined the following year in 2018 in their yearly returns. All but "ENPH" and "RUN" experienced a net negative yearly return, with "DQ" experiencing the worst loss with a 62.6% decrease in price. That is quite shocking, as the previous year had a net increase of over 200%. These drastic differences suggest that investing in "DQ" would be a risky choice. 

Please refer to the below-attached data table presenting our analysis of the stocks' performance in 2018.

After comparing the performances each year, our findings suggest that the best stock to invest in that could guarantee a high total daily volume and a positive yearly return is "ENPH". Over the two years of data we analyzed, it is clear that "ENPH" is consistently one of the most traded green energy stocks and increases their price, which is excellent news for a shareholder. 

### Original vs Refactored Code Performance 

Please refer to the below pop-up messages showing the elapsed run time for the original code.

Please refer to the below pop-up messages showing the elapsed run time for the refactored code.

It is evident that refactoring the original code improved the code's efficiency. The refactored code analyzed 2018 stock data 0.3516 seconds faster than the original code. The refactored code analyzed 2017 stock data 0.2812 seconds faster than the original code. By simplifying the script and introducing multiple arrays to reduce the number of loops, the refactored code can analyze large amounts of stock data with less time. 

## Summary

### Advantages and Disadvantages of Refactoring Code

Refactoring code can be very advantageous for analyzing large scripts of data. By simplifying the code without losing functionality, one can improve its efficiency because it takes less time for long scripts to run. In addition, by streamlining the code, its readability is improved, allowing users to scan through the code and understand the flow quickly. 

However, there are potential downsides to refactoring code. There is potential in altering the functionality if the user removes too much or alters the code's flow. Refactoring may involve introducing new variables and iterators. If the user is not consistent with the new additions, the user can run into syntax errors that may be difficult to debug.

These advantages applied when refactoring the original VBA script to analyze the stock data. The pop-up messages showing the elapsed run time indicate improved efficiency and shorter run-times associated with refactored code. In analyzing the original VBA script to analyze stock data, these disadvantages did not apply because both the original and refactored code results are the same. This finding means that the code's flow did not change after refactoring and that the functionality remained consistent.
