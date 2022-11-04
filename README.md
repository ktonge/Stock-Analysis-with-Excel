# Stock_Analysis

## Project Overview 

This module for class focused on our friend Steve seeking out more information about Green Energy Stocks.  This challenge specifically focuses on improving the VBA code I initially created for analyzing a short list of stocks by refactoring it.  This editing was intended to increase efficiency, by only looping through the data set once, rather than nested loops.  

I initially set out to find out information about stock DQ, and created macros to tell us information about just that one variable.  I then expanded that code outwards once using For loops to generate information about a list of stocks.  Learning how to refractor the data to create a usable spreadsheet that could easily be updatable with new information is the final step in this module.

## Results

### Initial Code

My initial VBA code used nested loops to find stock changes for a specified year in my data. 

![initialcode](https://github.com/ktonge/Stock_Analysis/blob/main/initialcode.png)

The main issue with this method was how long it took.  It was a quick method to put together, but its also repetatively looking through the loops over and over again to pull information

### Refactored Code

The refactored code was incredibly difficult for me,  I had to do alot of debugging and reformmating to get it to work.  It took alot more time to put together, however, the effects on efficiency was pretty incredible.   

![refactoredcode](https://github.com/ktonge/Stock_Analysis/blob/main/refactored_code1.png)
![refactoredcode](https://github.com/ktonge/Stock_Analysis/blob/main/refactored_code2.png)

A combination of three arrays store information, that way its going through the data in a much more organized way.  

##Summary

I think the benefits of refactoring data are best shown in the images below:


![runtime](https://github.com/ktonge/Stock_Analysis/blob/main/VBA_Challenge_2018.png.png)
![runtimerefactored](https://github.com/ktonge/Stock_Analysis/blob/main/VBA_Challenge_2018refactored.png.png)

Refactoring data reduced the time it took to ecexute from 1.83 to 0.28 seconds.   However, on the other side, the refactored code took me about five times the amount of time to write.  While the use of the array can be expanded more easily to look at for more possible stocks, if Steve was under a deadline to give the information to his client, the initial code does still work for its intended purpose.  
