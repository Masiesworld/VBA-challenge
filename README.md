# VBA-challenge
This is a project using VBA scripts to calculate and analyze stock market; it demonstrate the use of loops in VBA. 

# Instructions
## Part 1: 
** Create a script that will loop through all the stocks for one year and output the following information.**

The ticker symbol.

Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

The total stock volume of the stock.

You should also have conditional formatting that will highlight positive change in green and negative change in red.
## Part 2: 
Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". 
Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.
# Steps
1. **Define the variables.** Let's start to solve the problems on one worksheet first. It is clear to us that we need to use loops; therefore we need to determine the range of data we want to loop through. In this solution, variable *i* and *LastRow* are related to the row we are looping; variable *Summary_Table_Row* is related the row where we want to store the summary data in. *Ticker_total* stores the total volumes for each ticker. *start* and *change* are related to calculating the yearly change and percent change. 
2. **Print the headers.**
