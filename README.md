# Stock Analysis with VBA

## Overview of Project

### Purpose
The purpose of this data analysis was to analyze stock data to find the total daily volume and yearly return for each stock using refactored code.

## Results

### Stock Performance between 2017 and 2018
The code that was given created an InputBox that would run and analysis on the stock based on the year inputted by the user.
```
yearValue = InputBox("What year would you like to run the analysis on?")
```
The code then formatted the output by:
- having the value in cell A1 say "All Stocks (year inputted by the user)";
```
Range("A1").Value = "All Stocks (" + yearValue + ")"
```
- creating the header rows with "Ticker", "Total Daily Volume", and "Return";
```
Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"
```
- and initializing an array of the tickers.
```
Dim tickers(12) As String
    
tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"
```
I refactored the code to create a tickerIndex variable to access the correct index across four arrays: the ticker (as established above), tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
```
tickerIndex = 0
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```
