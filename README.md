# Green Stocks Analysis

## Project Overview

Steve's parents are passionate about green eneregies. They believe that as fossil fuels get used up there will be more reliance on alternative energy production. There are many forms of green energy to invest in, including hydro elctricity, wind energy, geothermal energy and bio energy. However, Steve's parents haven't done much research and have decided to invest all their money into DAQO New Energy Corporation, a company that makes silicon wafers for solar panels. Steve is concerned about diversifying their funds. He wants to analize a handful of green energy stocks in addition to DQO stock by creating an Excel file containing the information to be analyzed. Therefore, this project focuses on the analysis of the green stock data base to offer Steve's parent a complete overview of the which company is the best option to invest in. 

Initially, for the analysis of the data base we focused on obtaining the **total daily volume** of transactions and the **return percentage** of each company (by dividing the starting and ending prices), during 2017 and 2018. By obtaining that information we were able to compare the performance of each company by year, as well as highlighting whether the company's outcomes were negative or positive.. Although the code used to obtain the data runs correctly and delivers the expected results, the second part of the project focuses on editing, or refactoring, the code to make it faster and thus, in the future, be able to work with more information.

## Resources

* **Data source**: green_stocks.xlsm
* **Employed tools**: Visual Basic for Applications (VBA) in Excel

## Results 

### Stock performance between 2017 and 2018

The stock analysis outputs were presented in Excel as shown bellow: 

<img width="336" alt="Stocks_2017" src="https://user-images.githubusercontent.com/107893200/179088117-c878c59d-6bfa-4359-8a95-94d757447156.png">  <img width="335" alt="Stocks_2018" src="https://user-images.githubusercontent.com/107893200/179088136-9259466f-2d25-4919-a777-84488a93e18e.png">

Taking a closer look to the results we can conclude the following: 
- During 2017 the companies with better return percentage were DQ, ENPH, FSLR and SEDG.
- During 2018 the only two companies that had a positive return percentage were ENPH and RUN. 
- 






### Refactored script

In addition to generating code to get the information we needed to analyze which company might be a better option for Steve's parents to invest in, modifications were made to the code to make it run faster. This was intended to make the code work with a larger database. 

The **first part** of the code remained the same. The following steps were performed:
- Declare startTime an endTime as variables 
- Generate a message box to insert the year in which you would like to perform the analysis. 
- Start the timer to evaluate the amount of time the operation will take. 
- Format the output sheet on All Stocks Analysis worksheet
- Create a header row
- Initialize array of all tickers
- Activate the data worksheet
- Get the number of rows to loop over

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
  ```
The **second part** of the code is the one that contains the modifications to make it run faster:
- First, an Index containing each ticker, i.e. each company, was generated.
- Three output arrays were created: tickerVolumes, tickerStartingPrinces and tickerEndingPrices, which allowed to organize and store the data of each of the companies in a faster and cleaner way. 
- We create a for loop to initialize the tickerVolumes to zero. 
- We loop over all the rows in the spreadsheet.
- Each time the loop started the volume for the current ticker would be incremented.
- Then we create an if statement that will evaluate if the current row is the first row with the selected tickerIndex. 
- In the same way we create an if statement that will evaluate if the current row is the las row with the selected ticker, and if the if is not fulfilled the next step will be to...
- Increase the tickerIndex

```
1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
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

```
The **last section** arranges the information obtained in the Excel cells, formats the text and headings, and ends at the timer we activated at the beginning. 

```
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
```

### Execution times

The main change made to the code was the speed at which the analysis is performed. We can see this change by comparing the message box that pops up in each one of the versions. 

#### 2017

In the left image, the refactored code, we can see that the speed at which the code runs is 0.1054688. While in the image on the right, the initial code, the speed at which the code runs is 0.7421875.

<img width="422" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/107893200/179117304-9b660c5e-23a7-415e-896a-77ef21445c45.png"><img width="421" alt="2017" src="https://user-images.githubusercontent.com/107893200/179117323-1591fd48-6148-4b04-b28e-1d70a0b528b4.png">

#### 2018

In the left image, the refactored code, we can see that the speed at which the code runs is 0.109375. While in the image on the right, the initial code, the speed at which the code runs is 0.765625.

<img width="425" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/107893200/179117342-271312e1-57bf-47d0-9bad-53ae1275f94a.png"> <img width="424" alt="2018" src="https://user-images.githubusercontent.com/107893200/179117366-62172332-e202-462c-9a31-72100fd39251.png">


## Summary 

* What are the advantages or disadvantages of refactoring code?

One of the advantages of refactoring a code is the possibility of executing the same action many times in the same line of code, which makes it easier to read as well as to build the code itself. It becomes less time consuming and convoluted. One of the possible disadvantages of refactoring a code is the time it may take to generate the new code, as you may have errors in the beginning. 

In order to have better results when refactoring a code it is important to understand, prior to writing the code, the precise steps you want to carry out. This can help to identify if it is possible to join or resolve several steps in a single line or if it is possible to generate variables that can be reused.

* How do these pros and cons apply to refactoring the original VBA script?
