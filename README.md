# VBA_Challenge
To analyze stocks and through refactoring existing code reduce the amount of time it takes to perform that analysis
##	Overview of Project
      The purpose of this analysis is to provide a tool that can analyze large portions of the stock market using a simple excel sheet and the processing power of a standard PC.  While such large number crunching is left to larger server based calculation systems it may be possible to simplified code and restricted data parameters analyze large numbers of stocks quickly.
###	Results
    While analyzing more than 12 stocks is a daunting task it is possible to speed up the process and cover a years’ worth of data in under a full second. As can be seen in 
https://github.com/jrobertunder/VBA_Challenge/blob/main/resources/VBA_Challenge_2017.png
https://github.com/jrobertunder/VBA_Challenge/blob/main/resources/VBA_Challenge_2018.png
    Nesting for loops and parallel processing of the various variables streamlines the overall calculation process.
'1a) Create a ticker Index
        tickerindex = 0
        Sheets(YearValue).Activate

    '1b) Create three output arrays
   Dim totalvolume(12) As Double
   Dim tickerstartingprice(12) As Single
   Dim tickerendingprice(12) As Single
      
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        totalvolume(i) = 0
        tickerstartingprice(i) = 0
        tickerendingprice(i) = 0
        
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    
      
    For i = 2 To RowCount
        '3a) Increase volume for current ticker
        'add colume to corrosponding totalvolume
        totalvolume(tickerindex) = totalvolume(tickerindex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'if its the first row, save the price as startingprice
        If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            tickerstartingprice(tickerindex) = Cells(i, 3).Value
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        'If  Then
        'if its the last row, save the price as endigprice
        
        If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            tickerendingprice(tickerindex) = Cells(i, 6).Value
        '3d Increase the tickerIndex.
        
        tickerindex = tickerindex + 1
       
        End If
###	Summary: 

1.	What are the advantages or disadvantages of refactoring code?
	
    Refactoring code makes the programmer take a look at existing code and find ways to improve it. This however can lead to problems if the programmer tries to use pieces of old code that performs similar functions but cannot be made to fit into the existing code structure without completely rewriting it.  On the other hand reviewing existing code with an eye to streamlining the process can lead to leaps in programing that can dramatically improve the performance of the code.

2.	How do these pros and cons apply to refactoring the original VBA script?
    The original VBA stock analysis code took between .4 and .7 seconds to complete the operation depending on the year requested.  Refactoring the code, streamlining process and introducing parallel processing of the variables reduced that time to a uniform .1 second.  It is clear that refactoring the code to integrate improvements has made the process more efficient. 
