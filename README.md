# Module 2 | Assignment - Wall Street

Explore green energy stock performance by analyzing financial data using VBA.

## Overview of Project

### Purpose

The purpose of this project was to refactor the initial VBA macro AllSocksAnalysis in the new macro AllStocksRefactored for the file VBA_Challenge.xlsm. This process of refactoring had the goal of increasing the speed of the pattern created. The project acted as practice for showing understanding of VBA coding language and also the relationship of `for` loops, nested `for` loops, and conditionals for coding efficiency.

## Results

Through the help of groupwork, I was able to refactor the code to have the run times be cut down by up to 
.06 seconds. As seen in the below images:

![AllStocks_2017](https://github.com/drewabramo12/working_with_vba/blob/main/AllStocks_2017.PNG)
![AllStocksRefactor_2017](https://github.com/drewabramo12/working_with_vba/blob/main/AllStocksRefactor_2017.PNG)
![AllStocks_2018](https://github.com/drewabramo12/working_with_vba/blob/main/AllStocks_2018.PNG)
![AllStocksRefactor_2018](https://github.com/drewabramo12/working_with_vba/blob/main/AlStocksRefactor_2018.PNG)

The run time of 2017 and 2018 stock analyses went from 0.8632813 seconds down to 0.109375 seconds after refactoring. The change that occured was to remove the use of nested `for` loops for 3 separate `for` loops and the creation of output arrays. The creation of output arrays:
```
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```
were used to create individual variables that could save ticker values for a single `for` loop. The creation of the variable of `tickerIndex` also became valuable to use within the refactoring as it allowed for the repeated use of `tickerVolumes(tickerIndex)`, `tickerStartingPrices(tickerIndex)`, and `tickerEndingPrices(tickerIndex)`. These arrays meant more memory was used for storing variables but it also meant that fewer lines needed to be read as compared to the original AllStocksAnalysis macro. Another part of code that helped with the efficiency of the code is the conditional statement:
```
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value            
    '3d Increase the tickerIndex.
    tickerIndex = tickerIndex + 1
```
This conditional allows for the second `for` loop to run through the rows once. When the ticker values change in column A, the tickerIndex changes to address the new ticker value and the `for` loop can now use new variables in each of the same conditionals.

## Summary

- What are the advantages or disadvantages of refactoring code?

The advantages of refactoring code allows for a deeper understanding of the pattern of the initial code. It allows for the code to be manipulated into being more efficient, faster and possibly using fewer lines of code to run the same pattern. The disadvantages of refactoring depending on the complexity of the code can be both the time investment in changing the code and also the need to thoroughly update the code to apply and replace any variables that may have changed.

- How do these pros and cons apply to refactoring the original VBA script?

The pros and cons applied quite closely with the refactoring of the original VBA script. The lines of code became much faster and fewer iteration loops needed to occur. There was quite a bit of debugging that needed to occur for the newly refactored code to run correctly. Most of the bugs occured due to incorrect writing of new variables within the refactored code. This is something that should be alleviated with more practice on projects to come.
