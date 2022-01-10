# stock-analysis
## Overview of Project
      Steve would like to anaylse the stock martket to determine which stock(s) would be the best option for investment. With the initial VBA code, Steve was able to analyse certain stock information in the year of 2017 and 2018. The purpose of this project is to increase the efficiency of the initial VBA code so it can work with large set of stocks.

 ## Result

    Upon review of the data with the color code (green for positive return percentage and red for negative return percentage), it's clearly showed the majority of the 12 stocks that selected by Steve had positive return in the year of 2017. Particularly, DQ, SEDG, ENPH and FSLR have over 100% return in the year of 2017 and TERP was the only stock that had negative return. Furthermore, from the anaysis of the year 2018, although ENPH had positive return of 81.9%, it was 47.6% less than 2017. Surprisely RUN had 84% positive return when it only had 5.5% in the year of 2017. From this statistics, we can conclude that RUN is worth for investment.  
    
    ![VBA_Challenge_2017 (Refactored)](Resources/VBA_Challenge_2017 (Refactored).png)

    ![VBA_Challenge_2018 (Refactored)](Resources/VBA_Challenge_2018 (Refactored).png)

    To refactor the code, we have removed the lines that are in bold and italic below. After we have refactored the code, the analysis of the stock data did not change the result. It is still showing the same percentage of return for the selected stocks. we can see from the execution times of code has been significantly reduced from 2.29 seconds to 0.16 seconds for the year of 2017 and from 1.59 seconds to 0.25 seconds for the year of 2018. 
    
     ''2b) Loop over all the rows in the spreadsheet.
        Sheets(yearValue).Activate
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        'If Cells(j, 1).Value = tickers(tickerIndex) Then
            'tickerVolumes = tickerVolumes + Cells(j, 8).Value
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
            
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
        ***'And Cells(j, 1).Value = ticker(tickerIndex)***
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rows ticker doesnt match, increase the tickerIndex.
        ElseIf Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
        ***'And Cells(j, 1).Value = ticker(tickerIndex)***
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next j


## Summary

   Refactoring code can help us to have more organized and clear code that help to increase the efficiency of execution time and simplifed code for future updates and improvement. By refactoring code, user of the software and/or program will understand easily and be able to make update to the code without rewriting the brand new code for a similiar project. However, the disadvantage of refactoring code are it can be very time consuming and you may not know whether the refactored code really help to improve the quality and efficinency of the initial code. Addition, users might have to retest each of the execution code for their fuctionalities to avoid introduing new bugs to the code that may end up in dysfuction of the entire code. 
