# Analyzing Stocks with VBA
 
## Overview of Project

In this project, a VBA script that was used to deliver an analysis of a set of stocks with the click of a button is refactored in order for it to run more efficiently, using arrays and nested for loops.

## Results

A script was created that would loop through each stock ticker and output its total volume and return for a chosen year (either 2017 or 2018). A "ticker" variable was then created and assigned the value of the current stock ticker, and the total volume of each ticker was initialized to 0.

!["ticker" variable in the original code](/Resources/Old_Code_Variables.png)

Nested for loops and if statements were created that would check the value of the "ticker" variable, compare it to the current; previous; or next ticker in the spreadsheet, and increase the current ticker's totalVolume by the corresponding volume. These actions would only take place for items 0 through 11 in the "tickers" array, meaning that as soon as the loop reached the final ticker (VSLR) it would stop. 

![Loops in the original code](/Resources/Old_Code_Loops.png)

One of the most notable features of these loops is that they used two variables: i for the current ticker, and j for the current row in the spreadsheet. This was slightly inefficient because i was being assigned the index of the current ticker in every if statement. To refactor this code, we used i for the current row and gave the current ticker index its own variable, tickerIndex. We also turned totalVolume, tickerStartingPrices, and tickerEndingPrices into arrays, which would store the corresponding values for the current tickerIndex. The loops and if statements were changed accordingly.

![Variables and arrays in the refactored code](/Resources/Refactored_Code_Arrays.png)

![Loops in the refactored code](/Resources/Refactored_Code_Loops.png)

Since the variables in the original code were converted into arrays, the for loop to output the data was also adjusted to loop through each value in the arrays.

![Output loop in the original code](/Resources/Old_Code_Output.png)

![Output loop in the refactored code](/Resources/Refactored_Code_Output.png)

Below are the results of running the refactored code for both 2017 and 2018.

![Analysis results for 2017](/Resources/VBA_Challenge_2017.png)

![Analysis results for 2018](/Resources/VBA_Challenge_2018.png)

In 2017 the stock market performed very well, with almost all of the stocks receiving an increase in total return for the year. Two of the stocks, DQ and SEDG, even nearly doubled their price. However, 2018 saw the opposite outcome. A few of the stocks saw an increase in their total daily volume, but all except for ENPH and RUN saw a decrease in their yearly return, with DQ and JKS seeing more than 60% decreases. This indicates that maybe Steve's parents recommended he invest in DQ stock after looking at its 2017 returns, but they never looked at its 2018 returns.

![Popup message for the original 2017 analysis](/Resources/2017_Popup_Old.png)

![Popup message for the refactored 2017 analysis](/Resources/2017_Popup_Refactored.png)

![Popup message for the original 2018 analysis](/Resources/2018_Popup_Old.png)

![Popup message for the refactored 2018 analysis](/Resources/2018_Popup_Refactored.png)

Comparing these messages confirms that the refactored code runs more quickly for both years than the original code.

## Summary

The advantage of refactoring code is that you follow the DRY rule: Don't Repeat Yourself. Using one block of code (for example a for or while loop) to perform the same action multiple times can help declutter your code and make it run more efficiently. This can be helpful for analyzing very large data sets in a short amount of time. In this case, simplifying the loops to only use one variable in the conditional and using a single variable as the index for all of our arrays was more efficient than making the code re-check the value of i with each if statement. A disadvantage of refactoring code is that automating certain actions using nested for loops and variables can get confusing. You have to make sure that you're using the correct variable or array for every situation, and that all of your data types are matching up. I spent a few hours trying to figure out why the arrays in my refactored code were giving me "Subset out of range" errors, only to realize that I didn't need to give the tickerIndex variable a data type (like Single or Integer) and instead I could initialize it as "tickerIndex = 0".