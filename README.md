# Overview 
## Purpose
The client, Steve, originally requested a way to analyse a stocks performance in a given year by calculating Total Volume and Return. He is using this analysis for his parents to understand what areas might be good to invest in based on prior performance. Steve was able to see how well the stock performed and and then requested we expand the code to return analysis for a static set of tickers. Steve also found this analysis quite useful to compare across different stock performances. With this final refactoring, Steve has requested that we build an easy way for him to analyse a large dataset of stock data at the click of a button. This code will also analyse how efficent the performance is, so that Steve can see how long it takes for the analysis to be completed. The aim is to optimize the existing code to enhance performance time.

# Results

## 2017 vs 2018 Stock Performance
The comparison between each year stock performance is quite stark. In 2017, 11 out of 12 Stocks analysed have recorded a positive return varying between +5.5% (RUN) and +199.4% (CSIQ). Only 1 stock recorded a negative return of -7.2% (TERP) during this year. However, when we look at 2018 Return data, we can see that only 2 positive returns were recorded in this data set. The positive returns varied between +81.9% (ENPH) and + 84% (RUN), while the negative returns for 10 out of 12 stocks varied between -3.5% (VSLR) and -62.6% (DQ). 
This can allow us to conclude that generally 2017 was a much more profitable year for these stocks generally in comparison to 2018. The stocks which had the most deteriorated returns Year on Year are DQ and SEDG. The most improved returns were seen in RUN and ENPH.

<img src = "https://github.com/JerryMcG/stock_analysis/blob/main/Resources/2017_Results.png"> <img src = "https://github.com/JerryMcG/stock_analysis/blob/main/Resources/2018_Results.png">
## Refactoring the Code
### Original Code - Greenstocks.xlsm
Refactoring the code provided allowed us to optimize the code which is being run by our macros. The original code which was applied on the Greenstocks.xlsm file used:
1. An array for tickers which was defined within our code
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
2. Variables to hold each of the results calculated in our loops
```
Dim startPrice As Double
Dim endPrice As Double

For j = 2 To RowCount
   
      '- Find the total volume for teh current ticker
      If Cells(j, 1).Value = Ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value
      End If
      
      '- Find the starting price for the current ticker
      If Cells(j, 1).Value = Ticker And Cells(j - 1, 1).Value <> Ticker Then
        startPrice = Cells(j, 6).Value
      End If
      
      '- Find the ending price for the current ticker
      If Cells(j, 1).Value = Ticker And Cells(j + 1, 1).Value <> Ticker Then
         endPrice = Cells(j, 6).Value
      End If
    'close row loop
    Next j
```
3. Printed each Variable within the outer loop before going back to run all calculations again
```
Worksheets("All Stocks (" + yearValue + ")").Activate
    Cells(4 + i, 1).Value = Ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endPrice / startPrice - 1
```
4. Did not apply any formatting post results.

For 2017 & 2018 this code took approximately 0.95seconds to execute and return our results.
<img src = "https://github.com/JerryMcG/stock_analysis/blob/main/Resources/GreenStocksRuntime_2017.png" width = "400"> <img src = "https://github.com/JerryMcG/stock_analysis/blob/main/Resources/GreenStocksRuntime_2018.png" width = "400">

### Refactored Code: VBA_Challenge.xlsm
After refactoring, the code did execute much quicker on the VBA_Challenge.xlsm file despite adding some futher actions to complete:
1. Tickers are now gathered from sheet selected in Input and stored in array using a loop.
```
Dim tickers(12) As String

For j = 2 To RowCount
    'check if this is the first row with this ticker - if so, add to array
    If Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
        tickers(tickerIndex) = Cells(j, 1).Value
        tickerIndex = tickerIndex + 1
    End If
Next j
```
3. Created arrays to hold data for our calculations
```
Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

'adding data to arrays
For j = 2 To RowCount

    '3a) Increase volume for current ticker
    If Cells(j, 1).Value = Ticker Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
    End If
        
    '3b) Check if the current row is the first row with the selected tickerIndex.
    If Cells(j, 1).Value = Ticker And Cells(j - 1, 1).Value <> Ticker Then
        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
    End If
        
    '3c) check if the current row is the last row with the selected ticker
    'If the next rows ticker doesn't match, increase the tickerIndex.
    If Cells(j, 1).Value = Ticker And Cells(j + 1, 1).Value <> Ticker Then
        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
        '3d Increase the tickerIndex.
        tickerIndex = tickerIndex + 1
    End If
        
Next j
```
5. Using a new loop - print all data at the same time
```
For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
       
    Next i
```
7. Formatting being applied on results within another separate loop

These has significantly improved the efficiency of the script by completing the same tasks on all data before moving onto the next piece of the program. By storing the output of each of these loops in an array, we allow the values to be easily accessed by the code when it wants to leverage that data piece, rather than it only being temporarily held within a variable. The performance improved from approx 0.95seconds to 0.85seconds as you can see in the images below. Additionally, i found that my code did not run significantly faster by using a loop to populate the ticker array as opposed to defining it explicitly within the code. 

<img src = "https://github.com/JerryMcG/stock_analysis/blob/main/Resources/VBA_Challenge_2017.png"/> <img src = "https://github.com/JerryMcG/stock_analysis/blob/main/Resources/VBA_Challenge_2018.png"/>

# Summary
### What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code is that the basis of the code is there and existing functionality is operating as expected. Due to this we only need to write and append new code for the newly added functionality, it can make the enhancement of code much more simple. We can also reduce complexities within the code while further optimizing how the code operates. 
One of the disadvantages of refactoring code can come from how the previous creator has commented on what each part of the code is doing. This is essential to easily understand how the code operates before you begin to refactor. Sometimes code can seem overwhelming if it has not been well commented as the person attempting to refactor must take a lot of time to understand what the code is doing before attempting to refactor.This can cause code refactoring to be a very time consuming process. 

### How do these pros and cons appy to refactoring the original VBA Code?
This code was well commented and was clearly laid out with indendentation which meant it was easy to understand and easy to decihper where I needed to add some additional code. Reducing complexities within the code will directly impact the perfomance, which is one of the main measurements we have here of our successful code. 

For me, I found this refactoring quite time consuming because of one element that was required in the challenge. I spent a lot of time to figure out how to loop through each of the tickers in the selected sheet and then populate the tickers array based on that. I explored lots of different types of structures to store this data (Collections, Dictionaries) but found that i was not able to make much progress other than overwhelming myself further with added complexities - specifically how the array would only store unique values and not all tickers. 

Finally, with some help, i used a separate loop to populate tickers() before entering into loops for output arrays (tickerVolumes, tickerStartingPrices, tickerEndingPrices). I know there can be further optimization in my code by combining the following loops into one for loop: 
`For j = 2 To Rowcount`
However in light of the time already invested in creating the code as it is, I have decided not to refactor that item further at this time.

```
'1. tickers Loop
For j = 2 To RowCount
    'check if this is the first row with this ticker - if so, add to array
    If Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
     tickers(tickerIndex) = Cells(j, 1).Value
     tickerIndex = tickerIndex + 1
     End If
    Next j 

'2. tickerVolumes, tickerStartingPrices, tickerEndingPrices loops
      For j = 2 To RowCount

        '3a) Increase volume for current ticker
        If Cells(j, 1).Value = Ticker Then
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(j, 1).Value = Ticker And Cells(j - 1, 1).Value <> Ticker Then
         tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        If Cells(j, 1).Value = Ticker And Cells(j + 1, 1).Value <> Ticker Then
         tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
         '3d Increase the tickerIndex.
         tickerIndex = tickerIndex + 1
        End If
        
       Next j 
```
     
