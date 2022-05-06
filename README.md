# Stock Analysis - VBA Challenge

## Overview of Project
The purpose of this project was to analyze green energy stocks for a client so the client is able to recommend stocks based on the analysis. Analysis of green energy stocks was conducted to determine both the total daily volume and the return. The stocks were analyzed for both 2017 and 2018. 

To analyze the stocks, a VBA script was developed. However, the script only analyzed 12 stocks. If the client wanted to analyze a greater number of stocks, the VBA script developed for the 12 stocks may not work as well and may take longer to run. Therefore, the code was refactored so that a greater number of stocks could be analyzed quickly.  

## Results

### Refactored VBA Script
The VBA script was refactored to make the code more efficient. Originally, a nested for loop was used to move through each ticker array one stock at a time. The code was refactored to loop through the data one time and to collect all of the information rather than moving through each stock one at a time. 

Refactored Section of Code

```
   '1a) Create a ticker Index and set to zero
    tickerIndex = 0
    

    '1b) Create three output arrays for ticker (tickerVolumes as Long and tickerStartingPrices and tickerEndingPrices as single)
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
    
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
    
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If

        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

        End If
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            
            '3d Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If
            
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```



### Stock Performance: 2017 v. 2018
The overall return of investment for the green energy stocks in 2017 was positive whereas the overall return of investment for the green energy stocks in 2018 was negative. 

For 2017, there was only one stock with a negative return. TERP had a -7.2% return. The most successful return in 2017 was DQ with a 199.4% return followed by SEDG with a 184.5% return.

<img width="358" alt="Screen Shot 2022-05-05 at 6 12 59 PM" src="https://user-images.githubusercontent.com/103774401/167205090-c3cf3d1a-d8b8-4e35-a03f-3de1b49b0f9c.png">

For 2018, there were only two stocks with a positive return, which is a stark difference from 2017. ENPH had a positive return of 81.9% and RUN had a positive return of 84.0%. In 2017, ENPH had a return of 129.5%, so even though the return was still positive in 2018, the return value did drop. RUN, on the other hand, had a return of 5.5% in 2017 so its return increased significantly. The return for DQ dropped to -62.6% in 2018, and the return for SEDG was -7.8%. 

<img width="359" alt="Screen Shot 2022-05-05 at 6 13 44 PM" src="https://user-images.githubusercontent.com/103774401/167205131-e8c6f445-6c14-4377-af64-55d892cc55d0.png">

The differences between 2017 and 2018 show that stock returns can vary from year to year. This shows that before recommending a particular stock, it is good to examine how a stock's return does over time. 

### Execution Times: Original Script v. Refactored Script
After the code was refactored, the code ran significantly faster. When running the original code, the code ran in 0.62 seconds for 2017. For 2018, the original code ran in 0.64 seconds. 

<img width="564" alt="Green Stocks_2017" src="https://user-images.githubusercontent.com/103774401/167205193-070309bd-c8bc-4643-8b40-9fba411b0db0.png">

<img width="563" alt="Green Stocks_2018" src="https://user-images.githubusercontent.com/103774401/167205205-7f902e22-e99f-4c77-a7ee-ae4b25e2a2fb.png">

After refactoring the code, the code ran in 0.16 seconds for both 2017 and 2018. 

<img width="582" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/103774401/167205231-6ba3872a-2d54-4916-9250-7e7553fed6e4.png">

<img width="564" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/103774401/167205254-9876a917-1edd-4926-9189-ea0537561987.png">

The comparison of the execution times indicates that refactoring the original code led to the code running more quickly, so one of the goals of refactoring the code was successful. 

## Summary

Overall, refactoring the code for the VBA script was successful. It made the code more efficient and it ran much faster than the original code. 

### Advantages and Disadvantages of Refactoring Code
One advantage of refactoring code is that it makes the code more efficient. Additionally, refactoring the code makes it easier to read and understand for future readers and users of the code.

The largest disadvantage of refactoring code is that it can be time-consuming. If someone is developing code for a client, the client would have to decide if it is worth the time for the individual to refactor the code. Another disadvantage is that a mistake could be made which could lead to even more time spent refactoring the code. 

### Advantages and Disadvantages of Refactoring Stock Analysis VBA script.
The main advantage of refactoring the Stock Analysis VBA script was that it made the code more efficient, and it made the code run much faster than the original code. However, the script was only used to analyze 12 stocks. One disadvantage is that without looking at a much larger dataset, it is difficult to determine how much more efficient the refactored code is compared to the original code. Even though additional time was spent refactoring this code, it was still a good exercise as it made the code more efficient.
