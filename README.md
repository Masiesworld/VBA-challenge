# VBA-challenge
This is a project using VBA scripts to calculate and analyze stock market; it demonstrate the use of loops in VBA. 

# Instructions
## Part 1: 
**Create a script that will loop through all the stocks for one year and output the following information.**

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
2. **Print the headers.** This step is easier. Find the cells we are supposed to store the header information and insert values.
3. **Do for loop.** We know that we're going to loop through the whole worksheet and extract information by tickers. To avoid calculation error, first we need to discuss what if the opening price and/or closing price is zero. If they are non-zero, do calculations (as shown in the scripts) and print out the ticker name and the values we get. Noted that we reset the total volume back to zero and keep add one row each time encountering the end of one ticker. 
4. The other part is to **find the max values**. Note that it's outside the for loop in step 3. 
5. Next, we modify the codes so all these steps will be run on every worksheet. Specifically, we add *"ws."* in the cells and range related function and create **another for loop** outside the for loops in step 3. 
