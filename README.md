# VBA Challenge

## Overview of the project

The purpose of this analysis was to refactor code previously written so that it could be then applied to analysis of these and other stocks over a series of other years. The refactoring was able to cut down the processing time for the analysis so that if further information is added, it will not take as much time for the analysis to run. 
This analysis will also compare the 2017 stock performances for the stocks provided to the 2018 performances of those stocks.

### Results

It seems as though 2017 was overall a better performing year for this sector than 2018 was. As you can see, we coded the response so that the stocks that are highlighted in green show a positive performance and the ones that are coded in red show a negative performance. As shown below, there are only two stocks that performed well both years, which are **ENPH** and **RUN**. 


![VBA_Challenge_Chart_2017](https://user-images.githubusercontent.com/104734224/173445841-a15ab5c3-be5b-4d6e-b3ee-ef4697a6fe16.png)

![VBA_Challenge_Chart_2018](https://user-images.githubusercontent.com/104734224/173444921-5140557f-5ad2-43b5-b3ce-0392dce81834.png)

It is advisable to look at more current data to get a better analysis of which green stocks to invest in.

Now when it comes to processing time, refactoring the code definitely helped with that. As can be seen here, the original file ran in about .42 seconds for both years

![AllStocksTimer_2017](https://user-images.githubusercontent.com/104734224/173669523-6062d277-c2c7-494f-9be0-f9f6fc6fbb9f.png)

![AllStocksTimer_2018](https://user-images.githubusercontent.com/104734224/173669587-462354f0-06b0-4cfb-bcd8-673af19b9c5f.png)

After the refactoring, the processing time was about the same for 2017 and but was cut down to .039 in 2018. Some troubleshooting was attempted to run the data for 2017 but the processing time could not be cut down any further.  

![VBA_Challenge_2017](https://user-images.githubusercontent.com/104734224/173670366-f579fe1c-59da-4864-9af0-63ee94184f9d.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/104734224/173670433-edf18716-2843-4095-9374-a10b94790763.png)

The refactoring enabled the use of a index variable, which then could be applied to any number of stocks. As can be seen in the code below, the variable tickerIndex enables each calculated arrayed variable, instead of invidividually calculating each of them. 

'1a) Create a ticker Index
        tickerIndex = 0
        
    '1b) Create three output arrays
       Dim tickerVolumes(12) As Long
        
       Dim tickerStartPrices(12) As Single
        
       Dim tickerEndingPrices(12) As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        
             tickerVolumes(i) = 0
        
        Next i
            
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
            
        
        '3a) Increase volume for current ticker
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

The original code reads as the following:

        '4) Loop through tickers
                 
                 For i = 0 To 11
                 ticker = tickers(i)
                 totalVolume = 0
                 
         '5) loop through rows in the data
                 
                 Worksheets(yearValue).Activate
                 For j = 2 To RowCount
                 
          '5a) Get total volume for current ticker
                If Cells(j, 1).Value = ticker Then

                        totalVolume = totalVolume + Cells(j, 8).Value

So, instead of setting a for loop for each of the arrayed variables, it makes more sense to set an index and run the arrayed variables individually. 

Refactoring seems to have overall helped better organize the data and cut down processing time.

#### Summary

__1. What are the advantages or disadvantages of refactoring code?__

Refactoring code enables the benefit of learning different ways to accomplish the same taste. For instance, there was another set of code run that took longer in processing but was able to accomplish the same tasks in nested for loops. This was a bit inefficient because it ran through the main dataset 12 times over, instead of running each arrayed variable individually. It also allows better understanding of the code and it's application. Some disadvantages include the time involved troubleshooting refactored code, as one little misstep can set back the entire code, as well as knowing the limitations of the code being used. 

__2. How do these pros and cons apply to refactors the original VBA script?__

As stated, it was definitely a good learning opportunity to refactor the code in multiple ways. Additionally, a specific limitation with VBA is that with either code, the list of the initially defined variables must be in alphabetically order for the code to work. This presents some issues if working with thousands of different stocks. Futhermore, the code is not completely automated and must be adjusted to include additional information, such as more stocks. However, this code can run for the same set of stocks for multiple years. 
