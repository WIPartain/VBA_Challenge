# VBA_Challenge
Module 2 Project
## Overview
This project is an opportunity for our class to demonstrate our mastery of VBA and macros.  We utilized an assortment of stock price data over the span of two years, and we determined how effectively each stock performed in 2017 and 2018. 
### Methodology
The main concept of the project was to refactor the starter code provided to loop through a data set and generate an output determining the change in stock price for each ticker in 2 seperate years.  To do this we needed to do the following:
#### Step 1a:
Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.
#### Step 1b:
Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
The tickerVolumes array should be a Long data type.
The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.
#### Step 2a:
Create a for loop to initialize the tickerVolumes to zero.
#### Step 2b:
Create a for loop that will loop over all the rows in the spreadsheet.
#### Step 3a:
Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
Use the tickerIndex variable as the index.
#### Step 3b:
Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable.

![Code_pt1](https://github.com/WIPartain/VBA_Challenge/blob/main/Resources/Code_pt1.png)
#### Step 3c:
Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.
#### Step 3d:
Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.
#### Step 4:
Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.
![Code_pt2](https://github.com/WIPartain/VBA_Challenge/blob/main/Resources/Code_pt2.png)
## Results
The analysis of the stocks yielded the following results:
### 2017
![VBA_Results2017](https://github.com/WIPartain/VBA_Challenge/blob/main/Resources/VBA_Results_2017.png)
### 2018
![VBA_Results2018](https://github.com/WIPartain/VBA_Challenge/blob/main/Resources/VBA_Results_2018.png)

While all but one stock yielded positive results in 2017, the majority recorded a loss in 2018.  TERP was consistently down in both years while ENPH and RUN were both positive in 2017 and 2018.
 
### Refactored Solution
By refactoring the original code through the previously stated steps, we were able to greatly improve the speed of our algorithm. Despite the amount of rows in our data set, the refactored algorithm was able to generate results in one tenth of a second.
![Refactor2017](https://github.com/WIPartain/VBA_Challenge/blob/main/Resources/VBA_Challenge_2017.png)
![Refactor2018](https://github.com/WIPartain/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018.png)
## Summary
Refactoring code from a long series of conditionals helps to make it run much more efficiently.  You can potentially cut down on lines of code thus making it much simpler to organize, and it can run much faster which will aid any research conducted using that data set.  The only negative to refactoring is that it could take a considerable amount of time.  I ran into several issues trying to assemble the new array and the loops through the code however it was time well spent.  The code is more efficient, and it would be easy to make changes to the tickers and add additional rows of stock data.
