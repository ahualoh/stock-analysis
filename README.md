# stock-analysis
## BUT-VIRT-DATA-PT-10-2022-U-B-TTH | Module 2: VBA of Wall Street - Online Work

## Overview of Project: Explain the purpose of this analysis.
This analysis was done to evaluate the performance of selected green energy stocks. We use VBA to automate a macro so that an end use can quickly and simply, with the press of a button, run the analysis of the given data. The first iteration of this analysis was done specifically for the data sheets given - but in a second iteration, the code was refactored so that the macro coudl account for any additional cells if other data sheets are added to be analyzed. 

### The Data
The data we are given has two sheets, one holding data from 2017 and the other from 2018. There are a select number of stocks identified by their tickers. The data displays the performance of each ticker for every trading day in the year, with each row being a single day's performance data. The rows start with a ticker symbol, followed by the date, the prices throughout the day, and ends with the volume of trades from that day. The data is already sorted - first by ticker name, then by asecnding dates. 

### How to use the macro
To run the analysis, navigate to the the tab names "All Stocks Analysis", and press the button "Run Analysis" that is assigned to the macro. Once that button is pressed, you'll be asked to enter which year of data you'd like to analyze. After entering the year, the table output will display every unique ticker found in the respective data sheet, its associated accumulated volume, and it's change of return (as a percentage - using the very first closing price of the year, divded by the very last closing price of the year, minus 1). 

## Results: Comparing the stock performance between 2018 (original code) and 2018 (refactored code), as well as the execution times of the original script and the refactored script.

### Screenshots
The attached screenshots named "Code_Comparison_[]" show two windows - on the right is the original script, the left is the refactored script. There are boxes around lines of code that highlight where they are similar (with dotted lines) and different (solid lines). The colors of the boxes between two windows represent where the code is performing the same essential functions in each respective script. 

### Differences explained
The difference between the refactored code (0.1640625) and original code (0.8867188 seconds) is 0.8 seconds (0.7226563). Reference screenshots titled with "VBA_Challenge_2018". 

With the refactored code, we created a dimension for "tickerIndex" so that it can be applied as a variable across multiple arrays - tickers(), tickerVolumes(), tickerStartingPrices(), and tickerEndingPrices(), and we don't have to reference each of those array variables by their cell attributes. 

With Tickers being the first array, each variant listed in that array is a "key" to be matched in each row for the other respective information. *"Section 1a, Code_Comparison_1"*, of the refactored side assigns "tickerIndex" as the variable's position to match accroding to the Ticker array. Instead of `tickers` *[section 4, Code_Comparison_2]*,  `tickerIndex` *[section 4, Code_Comparison_2]*.

The code runs faster with these arrays because `tickerIndex` is now has a static position, so that we are only looping through the data once for each row. 


## Summary

### What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code are to simplify and make code more readable - resulting in it being easier to maintain. Refactored code *can* be faster and more efficient, and hopefully reduce room for error. 

The disadvantages of refactoring seem mainly to be time and expertise. If you don't have all the technical knowledge to know what and how to refactor, you could end up spending more time than it's worth, pending the program you're running. For a business especially time spent refactoring code can be costly. Refactoring could also introduce bugs, especially if you're not super sure what you're doing. 

### How do these pros and cons apply to refactoring the original VBA script?

For this particular script, the amount of lines in code didn't change by much, and most of the loops performed used the same formulas. Shaving off less than 1 second for this process is barely detectable. 

This code also had the advantage of being applied to "clean" data. There were no breaks in rows, everything was sorted according to their tickers, and by dates... is all the data was scrambled and not in order, the entirety of this code would have to be re-written. Which... instead of refactoring, I hear that sometimes it's better to start with a clean slate altogether. 