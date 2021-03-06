# Stock Analysis with VBA

## Overview of Project

### Purpose
The purpose of this data analysis was to analyze stock data to find the total daily volume and yearly return for each stock using refactored code in order to see which stocks performed best.

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
I created a for loop to initialize the tickerVolumes array to zero.
```
For i = 0 To 11
    tickerVolumes(i) = 0
Next i
```
I then created a loop over all the rows. Within the loop, the volume of the current tickerVolumes is increased with the tickerIndex variable as the index. The if-then statements check if the current row is the first row or the last row with the selected tickerIndex. If the current row is the first row, it is assigned the current closing price to the tickerStartingPrices variable. If the current row is the last row, it is assigned the current closing price to the tickerEndingPrices variable. The tickerIndex is then increased if the next row's ticker does not match the previous row's ticker.
```
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
         If Cells(i - 1, 1) <> Cells(i, 1) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         If Cells(i + 1, 1) <> Cells(i, 1) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
           
                '3d Increase the tickerIndex.
                tickerIndex = 1 + tickerIndex
            
            
        End If
    
    Next i
```
Finally, I created a for loop through the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays to output the "Ticker," "Total Daily Volume," and "Return" columns.
```
For i = 0 To 11
        
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
Next i
```
