# VBA_Stock_Analysis
<img width="660" alt="Screen Shot 2022-06-30 at 6 26 32 PM" src="https://user-images.githubusercontent.com/107282754/176789041-460c4441-39a7-49c7-ac28-5a73f166ff3d.png">

## Overview of the Project
In this challenge, we are going to refactor a specific client's dataset by using VBA to create a loop to loop through all the rows in order to collect an entire dataset at one time. This analysis will help determine if refactoring the original code successfully made the VBA script more efficient and run faster. 

## Purpose
This analysis will make the code more efficient and the stock dataset easier to read for future users.
## Analysis

#### Original Code

    Sub AllStocksAnalysis()

    yearValue = InputBox("What year would you like to run the analysis on?")

    Dim startTime As Single
    Dim endTime  As Single

       startTime = Timer
       
       '1) Format the output sheet on All Stocks Analysis worksheet
       Worksheets("All Stocks Analysis").Activate
       Range("A1").Value = "All Stocks (" + yearValue + ")"

       'Create a header row
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
    Dim startingPrice As Single
    Dim endingPrice As Single
      '3b) Activate data worksheet
        Worksheets(yearValue).Activate
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

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

Some changes were made to run the code faster and more efficent,in order to do that, I had to create different arrays such as tickerVolumes, tickerStartingPrices as well as tickerEndingPrices.

Sub AllStocksAnalysis()
#### Refactored Code

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
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim TickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
    
    ''2b) Loop over all the rows
    
    For i = 2 To RowCount
    
    '3a) Increase volume for current ticker
    
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

    '3b) Check if the current row is the first row with the selected tickerIndex.
    
     'If Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    TickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
    End If
        
    '3c) check if the current row is the last row with the selected ticker
    If Cells(i + 1, 2).Value <> tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 7).Value
                
    End If
            
    '3d Increase the tickerIndex.
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    tickerIndex = tickerIndex + 1
                
    End If
                
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(i + 4, 1).Value = tickers(i)
    Cells(i + 4, 2).Value = tickerVolumes(i)
    Cells(i + 4, 3).Value = tickerEndingPrices(i) / TickerStartingPrices(i) - 1
        
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
    
    For j = dataRowStart To dataRowEnd
    
    
        If Cells(j, 3) > 0 Then
            
            Cells(j, 3).Interior.Color = vbGreen
            
    Else
        
            Cells(j, 3).Interior.Color = vbRed
            
    End If
 
    Next j
    
    endTime = Timer
            MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
        
    End Sub
    
 ## Results of Code Performance
    
 2017: Original vs Refactored
   
   <img width="287" alt="All Stocks Analysis 2017 " src="https://user-images.githubusercontent.com/107282754/176970719-870ed771-8d7e-4ed7-aac0-cc2ec3217547.png">

<img width="290" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/107282754/176970742-61b73a96-b54e-4873-abc2-4a3305025217.png">

  

 2018: Original vs Refactored
   
   <img width="281" alt="All Stocks Analysis 2018" src="https://user-images.githubusercontent.com/107282754/176970690-2bdeef9f-f81c-48fd-be97-aa5414f85a7b.png">
  
   
   <img width="304" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/107282754/176970594-1b866d78-d0e1-428e-8b79-5356531fb212.png">
   
   After refactoring the codes, we were able to run the 2017 and 2018 faster 
   
#### Analysis of Stocks for 2017 and 2018

The conditonal formatting allows the stock perfomance for each selected company to be displayed and readable for the users

![Dataset Example 2017](https://user-images.githubusercontent.com/107282754/176971343-f579c04f-4dd9-46d0-ab0d-e98b91f096f5.png)

![Dataset Example 2018](https://user-images.githubusercontent.com/107282754/176971348-38c1c99a-42b1-4c19-beb2-89396e412d77.png)

### Summary

Refactoring makes a code more efficient and easy to read. As we have seen above, it reduces the processing time drastically and improves the code.



    

    
