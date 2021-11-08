# stock-analysis

## An Analysis of Stocks
Explore green energy stock performance by analyzing financial data using VBA.  

## Overview of Project
1. Create a VBA macro that can trigger pop-ups and inputs, read and change cell values, and format cells.
2. Use for loops and conditionals to direct logic flow.
3. Use nested for loops.
4. Apply coding skills such as syntax recollection, pattern recognition, problem decomposition, and debugging.

### Purpose
Steve's parents have decided to invest all of their money into DAQO New Energy Corp. 
We needed to find out if this was a worthy investment, or if there was another green energy company that would be more profitable.

We then refactored the code to run more efficiently.

## Results

### Comparing original analysis run times vs refactored
The original code ran in approximately .4 seconds for each year.

![Pic 1](https://github.com/charlagarcia/stock-analysis/blob/main/VBA_Challenge_2017.png)
![Pic 2](https://github.com/charlagarcia/stock-analysis/blob/main/VBA_Challenge_2018.png)


After refactoring the code, run times were significantly reduced to .09 seconds for 2017, and .05 seconds for 2018.

![Pic_3](https://github.com/charlagarcia/stock-analysis/blob/main/2017_Refactored.png)
![Pic_4](https://github.com/charlagarcia/stock-analysis/blob/main/2018_Refactored.png)

### How was the code refactored for optimization?
The original code contained a nested for loop.

![Pic 5](https://github.com/charlagarcia/stock-analysis/blob/main/VBA_Original_Code.png)

We made it more efficient by removing the nested loop, and removing the output to a separate loop.

![Pic_6](https://github.com/charlagarcia/stock-analysis/blob/main/VBA_Refactored_Code.png)
   
# Summary

## What are the advantages and disadvantages of refactoring code in general?
### Advantages
After refactoring, code is easier to understand and easier to maintain.

### Disadvantages
It can be time-consuming, and you might undo something that you have trouble fixing.

## What are the advantages and disadvantages of refactoring this particular VBA script?
### Advantages
Taking out our nested loop decreased run time significantly.

### Disadvantages
In order to refactor the code, testing has to be done with each change to make sure the code still works, as well as adding script to both the original and new code to find the run time of each in order to compare.



