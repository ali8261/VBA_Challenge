Purpose  

In this challenge , we are using VBA to run performance analysis on 12 stocks for 2017 and 2018 . After, running initial analysis, we are using refactoring strategy to find more efficient and faster way to analysis our data. 

Results
We created 4 arrays tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices 
1) 1st array ‘tickers’ was used for stock’s ticker symbol 
2) created ‘tickerIndex’ as variable to assign to other 3 tickers, and Will use this tickerIndex to access the correct index across the 3 different arrays on VBA Code.

 tickerIndex = 0
    For i = 0 To 11
    '5b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single, tickerEndingPrices(12) As Single
    
    '6a) Initialize ticker volumes to zero
    tickerVolumes(i) = 0
    
    Next i
    '6b) loop over all the rows
    
    For i = 2 To RowCount
    
        '7a) Increase volume for current ticker
       
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 9).Value
        
        '7b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 2).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 7).Value
            
            
        End If
        
        '7c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 2).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 7).Value
            

            '7d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '8) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i

Advantages:
1: makes our coding more efficient, and run faster  