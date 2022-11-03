# Stock Preformance Refactoring Analysis
## Project Overview
### Purpose and Background

The purpose of this analysis is to refactor the existing code to improve preformance and provide further insights to Steve as he continues his research. As Steve expands his data set we are providing a faster day to gather the necessary information on the "All Stocks Analysis" sheet. Steve will be able to use the refactored code whenever he adds more stocks to his analysis and for years to come.

## Stock Analysis Results

In  our initial code we use if statements to obtain the values on the 'All Stocks Analysis' sheet. While the code worked and provided us the correct information we needed, the if statements were checking the value of each row for the entire year of data. In our refactored code, we replaced the if statements with for loops, which decreased the time it took for the analysis to complete by .625 seconds as show below.

### Time for Analysis when using IF Statements
![Alt text](https://github.com/vstuopis/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png.png)

### Time for Analysis while using the for loops
![Alt text](https://github.com/vstuopis/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Refactored Code: An Advantage and Disadvantage
### Advantage:
One advantage of refactored code is the fact that by refactoring it, you are consolidating the code causing it to work moore efficiently. 

### Disadvantage:
One disadvantage of refactored code is that this newly refactored code may introduce new bugs into the code base. By touching the existing code that works, any changes could have downstream impacts which can be expensive to correct.

## VBA Scripts: Original v Refactored
### Original
One advantage of the original VBA script is the fact we were able to get the input and we know that the code works as written. One of the disadvantages of the original VBA script is that the code with all of the if statements needed to be read in it's entirety before moving on to the next row, making the code take longer to run.
### Refactored
An advantage of the refactored VBA script is that because of the for loops, the time it takes to run is much faster than the original script. One of the disadvantages of the refatcored VBA script is that with the time it took to refactor the code, the end result data wise remains the same so some may view this VBA script refactoring as a sunk cost.
