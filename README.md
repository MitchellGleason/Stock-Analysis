# Stock Analysis
## Overview of Project
#### The purpose of this project is to provide a visualization and analysis of green energy stocks. This will aid in decision making when investing in this sector, showing the total exchange volume and yearly return for each stock.
---
## Results
#### The outcome of this analysis shows that 2017 was overall a very good year for green energy stocks, with eleven out of the twelve selected stocks having a positive yearly return. More than just positive yearly returns, four of the twelve stocks reported a yearly return of more than 100%, with one reaching 199%. The total exchange volume between the stocks is relatively similar with only two outliers having much less and two outliers having much more exchange volume. The year 2018, however, has a very different result with only two of the twelve stocks having a positive yearly return. The total exchange volume is also more varied than 2017, with more stocks having high volume. The large discrepancy between the highest and lowest return stocks both years, as well as the difference between the number of stocks with a positive and negative yearly return, shows that this sector might be relatively volatile and could change very quickly.

This analysis was completed using Microsoft Excel with the addition of Visual Basic for Applications (VBA). In order to create a more robust code that would function with multiple worksheets and multiple year inputs. One of the first steps for this macro is to ask for a user input and assign it to a value which is later used to look through that specific years' worksheet.
```
yearValue = InputBox("What year would you like to run the analysis on?")
```
#### The first complicated part of the macro, after creating an array and assigning each stock to an index, is the acquisition of the number of rows for the user provided year. This is necessary as each year, and therefore each worksheet, will have a different number of rows depending on the stocks performance. This also allows for new data on new stocks to be taken without any change to the code. This was completed with the help of a [stackoverflow](https://stackoverflow.com/questions/18088729/row-count-where-data-exists) thread.
```
RowCount = Cells(Rows.Count, "A").End(xlUp).Row
```
#### The main body and function of this code is a nested For loop which gathers and adds all of the daily stock data together in order to view a yearly result of both return and exchange volume.
```
 'Create a ticker Index
    tickerIndex = 0

    'Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single

    'Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            'Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
```
#### Finally, startTime and endTime variables were added at the beginning, and the Timer statement was used to measure the time that the macro completed the analysis in.

![VBA_Channlenge_2017_NonReactored](https://user-images.githubusercontent.com/111290810/187256395-7c6965a6-2198-41f6-aac2-7e1638239b71.PNG)

![VBA_Channlenge_2018_NonReactored](https://user-images.githubusercontent.com/111290810/187256409-9a90b65c-feb8-4940-9aef-a45cc1dae4b3.PNG)

#### Seen above, the non refactored macro took about 0.74 seconds to complete. The final output results for this macro were simply printed within the For loop mentioned above. The refactored macro stored the gathered information in a set of arrays and output those arrays after the For loop was completed, resulting in much faster  execution times, seen here:

![VBA_Channlenge_2017](https://user-images.githubusercontent.com/111290810/187257207-2045e8c7-3546-4fb3-839c-15bd45934100.PNG)

![VBA_Channlenge_2018](https://user-images.githubusercontent.com/111290810/187257233-85d4d339-bc67-4c49-a649-753c5cd59494.PNG)

## Summary
#### The advantage of refactoring this code can be seen clearly above. While the data provided for these worksheets was relatively small, refactoring still gave a significant reduction in execution time, which would be even more noticeable with larger data sets. Additionally, refactoring can (and should) make the code more adaptable to new or different data, such as, in this case, more years of stock data. Refactoring can have disadvantages when the original code is perhaps too complicated to fully understand or simplify, especially if the person refactoring did not write the original code. It can also cause issues when the code is particularly large as it can create bugs that are more complicated than before. The refactoring done on this particular code, for example, did not include retrieving the stock ticker names, with this the macro could then accept any stock data and complete the analysis. However, this would require even more time spent on the code and is not needed to complete its intended purpose.
