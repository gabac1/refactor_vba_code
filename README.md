#Assessing the Performance of Different Stocks with VBA
##Overview of the Project
The goal of the project is to evaluate the performance of several stocks within the green energy industry and potentially attain some insight to help determine which of the stocks would be a good target to invest in. 12 stocks within the category are selected for this analysis, among which the DQ stock is given special attention.
Result
In 2017, all except one (TERP) of the 12 stocks yielded positive return, with DQ having the highest increase of 199.4%. However, in 2018, the majority of the stocks faced price drop of different extent, with DQ’s price decrease of 62.6% as the sharpest one. On the other hand, the only two stocks which maintained a growth in 2018, namely ENPH and RUN, rise significantly in their prices (81.9% and 84.0% respectively). These two stocks also had positive return in 2017.
A refactoring was done on the code used for this analysis. In the original code, to calculate the aggregated volume and yearly return of each stock, whether the ticker of the row is consistent with the stock ticker in selection would be checked when looping through each row, and at the end of each iteration, the calculated values are written directly to the result sheet:
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

The refactored code, on the other hand, does not check the ticker when calculating aggregated volume for each stock, rather, the ticker index is changed when the end of each stock’s data is reached, and the aggregating process would simply move on to the next stock. Also, the outputs of the loop are held in several arrays (tickers, tickerVolumes, tickerEndingPrices, tickerStartingPrices) before the final results on the sheet are derived from them:

    For i = 2 To RowCount
        ticker = tickers(tickerIndex)
    
        '3a) Increase volume for current ticker

        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)

        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = ticker And Cells(i - 1, 1).Value <> ticker Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value


            '3d Increase the tickerIndex.
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
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
    Next i
The refactored code runs significantly faster than the original one.
![alt text](https://github.com/gabac1/refactor_vba_code/blob/main/VBA_Challenge_2017.PNG)
For the original code, the time used is 0.4941406s.
![alt text](https://github.com/gabac1/refactor_vba_code/blob/main/VBA_Challenge_2018.PNG)
For the original code, the time used is 0.5039063s.
##Summary
In general, the refactoring process would make the original code to be more concise and efficient, but one bad thing is obviously that extra effort has to be made on changing an already working program. Also, by changing the design of the original code, some unexpected problems may occur when applying the code on different data.
For this project, the benefit of the refactoring is that the code runs faster, but it may also give rise to some potential problems.
For the original code of this project, the checking of start and end price would only work when the data is sorted by ticker and time. However, the method for the aggregation of volume would also work even on shuffled data. After refactoring, the code for volume would only work when the data is sorted as well. (In fact, the code can run even faster when ticker array is removed for the starting price and ending price search. Instead, simply checking for whether the previous or following ticker is the same as the current ticker would work:
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value


            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
) The original code might be less “clever”, but the straight forward logic of it can lead to higher adaptability under various scenarios.
