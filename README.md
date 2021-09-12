# stock-analysis

## Overview of Project

### Purpose 
 The purpose of this project was to facilitate Steves ability to perform analysis on a on a large data set over a period of 2 years (2017 & 2018).This was achieved by refactor a Microsoft Excel VBA code to collect certain stock information and determine whether or not the stocks are worth investing. Originally we wrote a code in similar format, however, the goal for this round was to increase the efficiency of the original code.
 
### Results

Below is a comparison of stock analysis for 2017 and 2018. The "Total Daily Volume", is summed over the course of the year. In the "Return" column, the stock's price at the end of the year is divided by the price at the beginning of the year, and converted to show percentage growth or loss. This shows how much return on investment a stock in a given corporation will provide the owner, with positive values marked ‘Green’ indicating increased value and negative values marked ‘Red” indicating losses.

With the help of conditional formatting we can see the differences between stock performance in 2017 and 2018 for the selected companies to be displayed easily and clearly. In 2017, only one ticker TERP shows a negative return of -7.21%. In 2018 most of the selected tickers showed a negative return. 
DQ did exceptionally well in 2017 with a 199.45% return however it dropped to almost -62.26. 

#### Stock Analysis 2017
![This is an image](https://github.com/ShahroseShahzad/stock-analysis/blob/main/Resources/Stock%202017.png)

#### Stock Analysis 2018
![This is an image](https://github.com/ShahroseShahzad/stock-analysis/blob/main/Resources/Stock%202018.png)

Code refactoring was a major part of this project. 
In order to improve the efficiency of my code, I created 3 new arrays: -tickerVolumes(12) to hold volume -tickerStartingPrices(12) to hold starting price -tickerEndingPrices(12) to hold ending price

In these arrays performance data was stored for each stock when a for loop runs Macro analysis on them. In the original code the tickers array establishes a ticker symbol that can be called on for each stock.
Matching the 3 performance arrays with the ticker array is done by using a variable called the tickerIndex. and setting the tricker index to 0.

After creating the arrays I used Nested For Loops and variables to loop through the data and complete the analysis.
An example of the Refactored Vs Original code is shown below:


#### Refactored Code
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
```
#### Original Code 

Sub Macrocheck()

Dim textMessage As String

testMessage = "Hello World! "

MsgBox (testMessage)

End Sub

Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

        'Create a header row
        Cells(3, 1).Value = "Year"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate

        'set initial volume to zero
        totalVolume = 0

            Dim startingPrice As Double
            Dim endingPrice As Double

                 'find the number of rows to loop over
                 RowCount = Cells(Rows.Count, "A").End(xlUp).Row

                'loop over all the rows
                 For i = 2 To RowCount

                      If Cells(i, 1).Value = "DQ" Then

                         'increase totalVolume by the value in the current row
                          totalVolume = totalVolume + Cells(i, 8).Value

                      End If

                      If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

                        startingPrice = Cells(i, 6).Value

                      End If

                      If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

                        endingPrice = Cells(i, 6).Value

                      End If

               Next i

    Worksheets("DQ Analysis").Activate
    
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1


End Sub

Sub AllSockAnalysis()

   '1) Format the output sheet on All Stock Analysis work Sheet
   
      Worksheets("All Stocks Analysis").Activate
      
    
       'Title Cell A1
        Range("A1").Value = "All Stocks (2018)"
        
        'Create aheader row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

    '2) Initialize array of all tickers
       
          Dim tickers(11) As String
            
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

    '3a) Initialize variables for starting price and ending price
  
        Dim startingPrice As Double
        Dim endingPrice As Double
        
    '3b) Activate data worksheet
    
        Worksheets("2018").Activate
        
     '3c) Get the Numbers of the rows to loop over
        
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
    '4) Loop through tickers
    
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
            
            '5) loop thourhgrows in the data
            Worksheets("2018").Activate
           
             For j = 2 To RowCount
            
                    '5a) Get total volume for current ticker
                    If Cells(j, 1).Value = ticker Then
                        
                        totalVolume = totalVolume + Cells(j, 8).Value
                        
                    End If
                        
                    '5b) Get starting price for current ticker
                    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    
                        startingPrice = Cells(j, 6).Value
                        
                    End If
                    
                    '5c) Getending price for current ticker
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
        
   End Sub
   
   Sub formatAllStockAnalysisTable()
    
        Dim startTime As Single
        Dim endTime  As Single

            yearValue = InputBox("What year would you like to run the analysis on?")

        startTime = Timer
   
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

                              'Color the cell green
                              Cells(i, 3).Interior.Color = vbGreen

                         ElseIf Cells(i, 3) < 0 Then

                              'Color the cell red
                               Cells(i, 3).Interior.Color = vbRed

                         Else

                            'Clear the cell color
                             Cells(i, 3).Interior.Color = xlNone

                        End If

                 Next i
                 
                          endTime = Timer
            
                            MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

   End Sub
 
   Sub ClearWorksheet()
   

            Cells.Clear
   
   End Sub
   
   Sub yearValueAnalysis()
   
        Sheets("All Stocks Analysis").Activate
        yearValue = InputBox("What year would you like to run the analysis on?")
        Range("A1").Value = "All Stocks (2018)"
        Range("A1").Value = "All Stocks (" + yearValue + ")"
    
   End Sub
Sub Output()
    Worksheets("Output").Activate
    
    
  ' Make a list of square numbers
    For i = 1 To 10

    Cells(1, i).Value = i * i

    Next i

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
