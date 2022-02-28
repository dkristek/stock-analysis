# Stock Analysis Using VBA 
## Overview of Analysis
I was tasked with refactoring the VBA code of an Excel Spreadsheet that display the performance of various stocks over the years of 2017 and 2018. The original code, which analyzes 12 different stock tickers, was adapted from a code that only analyzed one stock ticker over a one-year period. This led to a somewhat unoptimized code. The goal of this refactoring was to produce a faster, more efficient code than the provided code. 

The dataset, which was provided in an .xlsx format, contains two tables ‘7’ and ‘2018’. Each table contains the daily stock market data (open price, high and low, close price, adjusted close price, and volume) on 12 different stock tickers. The code analyzes the stocks by finding the close price of the stock on the first date of the year and the close price (starting price) on the last day of the year (ending price) and dividing the ending price by the starting price and subtracting 1 from the quotient. The refactored code should produce the same results as the old code but faster. 

## Results
The original code operated by used nested for loops to complete the analysis of the stocks. An array of the twelve stock tickers was created and initialized. The first for loop would loop through the ticker array and a nested for loop would loop through the entirety of the stock dataset to collect the needed data. The script would print the data to the spreadsheet and move to the next ticker in the array. Thus, the script would run through the whole dataset a total of twelve times to complete the analysis. The original code can be found in the collapsible section below.
<details><summary>Original Code</summary>
<p> Here is the original VBA code
  
 ```
  Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime  As Single

    Worksheets("All Stocks Analysis").Activate
    
    'get year for analysis
    yearValue = InputBox("What year would you like to run the analysis on?")
        startTime = Timer
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create Title and headers
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Cells(3, 4).Value = "Return Percentage"
    'initialize array of all tickers
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
    
    'initialize start and end price
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'activate worksheet containing data
    Sheets(yearValue).Activate
    
    'find number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'loop through ticker
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        'loop through rows in the data
        Sheets(yearValue).Activate
        For j = 2 To RowCount
            'find total vol for cur ticker
            If Cells(j, 1).Value = ticker Then

                totalVolume = totalVolume + Cells(j, 8).Value
    
            End If
            
            'find start price for cur ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
                startingPrice = Cells(j, 6).Value
            
            End If
            
            'find end price for cur ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1) = ticker Then
            
                endingPrice = Cells(j, 6).Value
            
            End If
            
        Next j
        
        'output data for current ticker
        Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Cells(4 + i, 4).Value = (Cells(4 + i, 3).Value * 100)
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:D3").Font.Bold = True
    Range("A3:D3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:D3").Font.FontStyle = "Bold Italic"
    Range("A3:D3").Font.Size = 12
    Range("A3:D3").Font.Color = RGB(0, 0, 255)
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.000"
    Range("D4:D15").NumberFormat = "0.00\%"
    Columns("A:D").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    
  
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            'Change cell color to green
            Cells(i, 3).Interior.Color = vbGreen
            
        ElseIf Cells(i, 3) < 0 Then
        
            'Change cell color to red
            Cells(i, 3).Interior.Color = vbRed
            
            
        End If
        
    Next i
    
        For i = dataRowStart To dataRowEnd
        
        If Cells(i, 4) > 0 Then
            
            'Change cell color to green
            Cells(i, 4).Interior.Color = vbGreen
            
        ElseIf Cells(i, 4) < 0 Then
        
            'Change cell color to red
            Cells(i, 4).Interior.Color = vbRed
            
            
        End If
        
    Next i
  
        endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```  
  </p>
</details>
The results of the analyses using the original code is shown below.

![Analysis Results for the year 2017 using the old code](https://github.com/dkristek/stock-analysis/blob/main/Resources/2017_StocksOld.png)![Analysis Results for the year 2018 using the old code](https://github.com/dkristek/stock-analysis/blob/main/Resources/2018_StocksOld.png)

The aim of the refactoring process was to decrease the runtime of the script. The original script looped through the dataset 12 times (the length of the ticker array), thus making it a good target for refactoring. Instead of using nested loops the refactored script will loop through the entire dataset only once, collecting all the needed data in one pass. In the refactored code, arrays were created for the tickers, the ticker volume, and the starting and ending prices. A ticker index was created to keep track of the stock ticker and the for loop would increase the ticker index when it detected that the ticker had changed on the dataset. In the for loop, ticker index was used to access the stock ticker index of the four arrays. This allowed the for loop to read and store values for all the arrays while looping through the stock dataset. After the for loop had finished, a new for loop was created which printed the data stored in the arrays to a spreadsheet. The full refactored VBA code can be found below.
<details><summary>Refactored Code</summary>
<p> Here is the full refactored VBA code
  
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
    
    'Create a ticker Index
    tickerIndex = 0
    
    'Create three output arrays
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single
    
    ''For loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0

    Next i


  'Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

        'verify value is correct ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
    
            'Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
    
        'Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    
        'set starting price
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
    
        'Check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
            
            'set ending price
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
            'increase ticker index
            tickerIndex = tickerIndex + 1
          
        End If
        
    Next i


    'Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
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

        'conditional formatting green for + return red for - return
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
``` 
  </p>
</details>

The analysis results using the refactored code are shown below.

![Results for the year 2017 using the refactored code](https://github.com/dkristek/stock-analysis/blob/main/Resources/2017_stocks.png) 
![Results for the year 2018 jusing the refactored code](https://github.com/dkristek/stock-analysis/blob/main/Resources/2018_stock.png)

The original code had a runtime of 0.8125 seconds for the year of 2018 and 0.7578 seconds for the year of 2017. While the refactored code had a runtime 0.0938 seconds for the year of 2018 and 0.0938 seconds for the year of 2017. These results indicate that the refactored code is significantly faster than the original code. The speeds for the original code are shown below. 

![Speed for 2018 using original code](https://github.com/dkristek/stock-analysis/blob/main/Resources/2018_SpeedOld.png)
![Speed for 2017 using the original code](https://github.com/dkristek/stock-analysis/blob/main/Resources/2017_speedOld.png)


The runtimes for the refactored code are shown below.

![Speed for 2018 with refactored code](https://github.com/dkristek/stock-analysis/blob/main/Resources/2018_speed.png)
![Speed for 2017 with refactored code](https://github.com/dkristek/stock-analysis/blob/main/Resources/2017_Speed.png)

## Summary
Refactoring is the process of restructuring and reorganizing pre-existing code. Common objectives of refactoring code are to increase readability, simplify needlessly complex code, and to increase performance and efficiency. The main benefits of refactoring are the improvement of the code’s design, which in turn can improve the performance and the capability of the code; and making the code easier to read and understand which can make finding bugs a simpler and faster process.  While there are many benefits to refactoring code it is not without risk. It is possible to introduce more bugs into the code during refactoring which could leave the code in a worse state than previously. Another risk of refactoring relates to the fact that it is hard to objectively define what ‘clean’ or ‘neat’ code looks like. If a clear objective is not defined and communicated to the refactoring team it is possible to end up with a code that is not more efficient or easier to maintain than the original code. This would result in a waste of time and money with no obvious benefit while increasing the chance of introducing new bugs into your program. 

While the refactored code ran faster, there are advantages and disadvantages to both forms of the code. The original code is slower as it uses nested for loops that cause the code to run more slowly. However, the original code is arguably easier to understand conceptually. The refactored code creates an index that is used to access several different arrays to store data while only looping through the whole dataset once. It is conceivable that some might find this concept more difficult to understand than nested for loops. The advantages of the refactored code are its speed and conciseness.
