# Module 2 Challenge
 
## Overview of Project

In Module 2, we created scripts in Microsoft VBA to help Steve analyze a set of stocks with the click of a button. In this challenge, we refactor the code from Module 2 in order for it to run more efficiently, using arrays and nested for loops.

## Results

![Analysis results for 2017](/Resources/2017_Results.png)

![Analysis results for 2018](/Resources/2018_Results.png)

In 2017, the stock market performed very well, with almost all of the stocks receiving an increase in total return for the year. Two of the stocks, DQ and SEDG, even nearly doubled their price. However, 2018 saw the opposite outcome. A few of the stocks saw an increase in their total daily volume, but all except for ENPH and RUN saw a decrease in their yearly return, with DQ and JKS seeing more than 60% decreases. This indicates that maybe Steve's parents recommended he invest in DQ stock after looking at its 2017 returns, but they never looked at its 2018 returns.

![Popup message for the original 2017 analysis](/Resources/2017_Popup_Old.png)

![Popup message for the refactored 2017 analysis](/Resources/2017_Popup_Refactored.png)

![Popup message for the original 2018 analysis](/Resources/2018_Popup_Old.png)

![Popup message for the refactored 2018 analysis](/Resources/2018_Popup_Refactored.png)

Compared to the original script, the refactored script runs noticeably faster.

##Summary

The advantage of refactoring code is that you follow the DRY rule: Don't Repeat Yourself. Using one block of code (for example a for or while loop) to perform the same action multiple times can help declutter your code and make it run more efficiently. This can be helpful for analyzing very large data sets in a short amount of time, like what we did in this analysis. A disadvantage, though, is that automating certain actions using nested for loops and variables can get confusing. You have to make sure that you're using the correct variable or array for every situation, and that all of your data types are matching up. I spent a few hours trying to figure out why the arrays in my refactored code were giving me "Subset out of range" errors, only to realize that I didn't need to give the tickerIndex variable a data type (like Single or Integer) and instead I could initialize it as "tickerIndex = 0".