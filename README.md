# stock-analysis

## Overview of Project

### Purpose 
 The purpose of this project was to refactor a Microsoft Excel VBA code to collect certain stock information in the year 2017 and 2018 and determine whether or not the stocks are worth investing. This process was originally completed in a similar format, however, the goal for this round was to increase the efficiency of the original code.
### Results
Before refactoring the code, I began by copying the code that was needed to create the input box, chart headers, ticker array, and to activate the appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring. Below is the instruction and code as written in the file.

```
Sub AllStocksAnalysisRefactored()
    
        Dim startTime As Single
        Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
             TickerIndex = 0
             
    '1b) Create three output arrays
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
                
                tickerVolumes(TickerIndex) = tickerVolumes(TickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
                             If Cells(i, 1).Value = tickers(TickerIndex) And Cells(i - 1, 1).Value <> tickers(TickerIndex) Then
                                 
                                     tickerStartingPrices(TickerIndex) = Cells(i, 6).Value
                             
                             End If
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
                             If Cells(i, 1).Value = tickers(TickerIndex) And Cells(i + 1, 1).Value <> tickers(TickerIndex) Then
                                
                                  tickerEndingPrices(TickerIndex) = Cells(i, 6).Value
                        
                            End If
            

            '3d Increase the tickerIndex.
            
                      If Cells(i, 1).Value = tickers(TickerIndex) And Cells(i + 1, 1).Value <> tickers(TickerIndex) Then
                                        
                                        TickerIndex = TickerIndex + 1
                     End If
            
        'End If
    
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
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

Sub Clear()

    Cells.Clear

End Sub

```



By refactoring the code I was able to reduce my Macro Run Time for both 2017 as well as 2018.
#### Original Run time 2017
![This is an image](https://github.com/ShahroseShahzad/stock-analysis/blob/main/Resources/Original%202017.png)

#### Refactored Run time 2017
![This is an image](https://github.com/ShahroseShahzad/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

#### Original Run time 2018
![This is an image](https://github.com/ShahroseShahzad/stock-analysis/blob/main/Resources/Original%202018.png)

#### Refactored Run time 2018
![This is an image](https://github.com/ShahroseShahzad/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### Advantages and Disadvantages of refactoring code 
Refactoring helps us to improve the internal structure of an existing source code, while maintaining its external behavior. 

Refactoring has many advantages that helps to make our code more organized and save time. 
Once refactoring has taken place, the code easier to understand and read, as well as easier to maintain. A cleaner code includes design and software improvement, debugging, and faster programming. It is also beneficial for other users who view our projects because it becomes easier to read, as the code is more concise and straightforward. If a code is not properly written and a lot of bugs are raised. Before fixing bugs code should be refactored. 


However, refactoring does have a few disadvantages. It can be a very time-consuming process and should not be done close to a deadline. The cost of refactoring is a lot higher than if the code was written from scratch. If you do not test the refactoring code and have enough time to fix it, it can introduce bugs causing more delays.


### Advantages and Disadvantages of refactoring code in Original VS Refactored Script
One if the advantages of refactoring for me was a decrease in macro run time. The original analysis took approximately 0.156seconds whereas our new analysis only took about a four of the time (approximately 0.25 seconds) to run. Attached below are the screenshots that indicate the run time for our new analysis.


![This is an image](https://github.com/ShahroseShahzad/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![This is an image](https://github.com/ShahroseShahzad/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
