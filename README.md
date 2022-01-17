# Stock-Analysis Using VBA

VBA of Wall Street

# Overview

1. Create a VBA macro that can trigger pop-ups and inputs, read and change cell values, and format cells.
2. Use for loops and conditionals for the logic flow.
3. Use nested for loops.
4. syntax recollection, pattern recognition, problem decomposition, and debugging.

## Purpose
As part of an assignment for the Boot Camp, an initial stock analysis was conducted for Steve. The original code was then refactored to loop only once.  The purpose is to determine if the refactored changes made an impact on the run time.  

Using run buttons, Steve will have the ability to put in the year into an input box, which removes any magic numbers.

# Results
## Comparing 2017 and 2018 Stock Analysis
Comparing the 2017 and the 2018 stocks, the difference in the total daily volume between the two years that resulted in less than a $100,000,000 in increased volume was not enough to generate a positive 2018 return percentage.  The tickers ENPH and RUN had would have been considered good investments due to the positive returns in 2018, and both of the tickers had increases greater than $200,000,000 over the 2017 total daily volumes. 
 
![Pic 1](https://github.com/josepcherian/green_stocks_Challenge/blob/main/1Compare_2017_2018.PNG)

## Comparing the Original Run Times to the Refactored Run Times

Run times for the original code took around .4 seconds.
![Pic 2](https://github.com/josepcherian/green_stocks_Challenge/blob/main/2%20Original_2017%20for%200.4%20sec.png)![Pic 3](https://github.com/josepcherian/green_stocks_Challenge/blob/main/3%20Original_2018%20for%200.4%20sec.png)


Run times for the refactored code took around .08 seconds, which is displayed in scientific notation: 8 E -02.
![Pic 4](https://github.com/josepcherian/green_stocks_Challenge/blob/main/4%202017_Refactored%20for%200.08%20sec.png)![Pic 5][https://github.com/josepcherian/green_stocks_Challenge/blob/main/5%202018_Refactored%20for%200.08%20sec.png)


Refactoring the code did make the run times decrease, which optimizes the code. 

## What was the difference in the codes? 

The original code contained a nested for loop.  Included in that outer loop, it also output the data for the current data, results in a significant number of iterations.  


With refactoring the code to only have one loop and moving the output to a separate loop, the number of iterations was reduced.



# Summary
## 1.Advantages or disadvantages of refactoring code
### Advantages
Finding the root cause of a potential bug can be done with refactoring. Anyone can review and come with a fresh perspective and can tidy up the code to make it run leaner. 

### Disadvantages

Programming can have multiple approaches to a solve a problem.  In those differences, programmers may have alternative logic steps, which will require testing to see how those differences play out in the script.  Refracting a stable code to apply a different set of logic could be costly or introduce new bugs into the system.  
	
## 2. How do these pros and cons apply to refactoring the original VBA script?
Reducing the number of loops decreases the memory needed for processing the data, which reduces the run time and optimizes the performance of the script. 
To refactor the code, testing has to be done with each new addition to check for the efficiency of the new code.  
