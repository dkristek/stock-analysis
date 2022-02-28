# Stock Analysis Using VBA 
## Overview of Analysis
I was tasked with refactoring the VBA code of an Excel Spreadsheet that display the performance of various stocks over the years of 2017 and 2018. The original code, which analyzes 12 different stock tickers, was adapted from a code that only analyzed one stock ticker over a one-year period. This led to a somewhat unoptimized code. The goal of this refactoring was to produce a faster, more efficient code than the provided code. 

The dataset, which was provided in an .xlsx format, contains two tables ‘7’ and ‘2018’. Each table contains the daily stock market data (open price, high and low, close price, adjusted close price, and volume) on 12 different stock tickers. The code analyzes the stocks by finding the close price of the stock on the first date of the year and the close price (starting price) on the last day of the year (ending price) and dividing the ending price by the starting price and subtracting 1 from the quotient. The refactored code should produce the same results as the old code but faster. 

## Results
The original code operated by used nested for loops to complete the analysis of the stocks. An array of the twelve stock tickers was created and initialized. The first for loop would loop through the ticker array and a nested for loop would loop through the entirety of the stock dataset to collect the needed data. The script would print the data to the spreadsheet and move to the next ticker in the array. Thus, the script would run through the whole dataset a total of twelve times to complete the analysis. [PIC OF CODE] 
The results of the analyses using the original code is shown below.

![Analysis Results for the year 2017 using the old code](https://github.com/dkristek/stock-analysis/blob/main/Resources/2017_StocksOld.png)![Analysis Results for the year 2018 using the old code](https://github.com/dkristek/stock-analysis/blob/main/Resources/2018_StocksOld.png)

The aim of the refactoring process was to decrease the runtime of the script. The original script looped through the dataset 12 times (the length of the ticker array), thus making it a good target for refactoring. Instead of using nested loops the refactored script will loop through the entire dataset only once, collecting all the needed data in one pass. In the refactored code, arrays were created for the tickers, the ticker volume, and the starting and ending prices. A ticker index was created to keep track of the stock ticker and the for loop would increase the ticker index when it detected that the ticker had changed on the dataset. In the for loop, ticker index was used to access the stock ticker index of the four arrays. This allowed the for loop to read and store values for all the arrays while looping through the stock dataset. After the for loop had finished, a new for loop was created which printed the data stored in the arrays to a spreadsheet. [excerpt of script] 
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
