# Overview 
## Purpose
The client, Steve, originally requested a way to analyse a stocks performance in a given year by calculating Total Volume and Return. He is using this analysis for his parents to understand what areas might be good to invest in based on prior performance. Steve was able to see how well the stock performed and and then requested we expand the code to return analysis for a static set of tickers. Steve also found this analysis quite useful to compare across different stock performances. With this final refactoring, Steve has requested that we build an easy way for him to analyse a large dataset of stock data at the click of a button. This code will also analyse how efficent the performance is, so that Steve can see how long it takes for the analysis to be completed. The aim is to optimize the existing code to enhance performance time.

# Results

using code and images and headesr for different sections

Both hardcoded lists are quicker by 0.01sec - but dynamic list is more scalable.

# Summary
### What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code is that the basis of the code is there and existing functionality is operating as expected. Due to this we only need to write and append new code for the newly added functionality, it can make the enhancement of code much more simple. We can also reduce complexities within the code while further optimizing how the code operates. 
One of the disadvantages of refactoring code can come from how the previous creator has commented on what each part of the code is doing. This is essential to easily understand how the code operates before you begin to refactor. Sometimes code can seem overwhelming if it has not been well commented as the person attempting to refactor must take a lot of time to understand what the code is doing before attempting to refactor.This can cause code refactoring to be a very time consuming process. 

### How do these pros and cons appy to refactoring the original VBA Code?
This code was well commented and was clearly laid out with indendentation which meant it was easy to understand and easy to decihper where I needed to add some additional code. Reducing complexities within the code will directly impact the perfomance, which is one of the main measurements we have here of our successful code. 

For me, I found this refactoring quite time consuming because of one element that was required in the challenge. I spent a lot of time to figure out how to loop through each of the tickers in the selected sheet and then populate the tickers array based on that. I explored lots of different types of structures to store this data (Collections, Dictionaries) but found that i was not able to make much progress other than overwhelming myself further with added complexities - specifically how the array would only store unique values and not all tickers. 

Finally, with some help, i used a separate loop to populate tickers() before entering into loops for output arrays (tickerVolumes, tickerStartingPrices, tickerEndingPrices). I know there can be further optimization in my code by combining the following loops into one for loop: 
`For j = 2 To Rowcount`

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
     
