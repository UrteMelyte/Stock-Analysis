# Stock-Analysis

## Overview of Project
The purpose of this project is to take a previous VBA script and refactor it so that it loops through the data one time, with the intent of making the script run faster. We are currently analyzing two years of stock data for 12 tickers, but if we were to expand our data set our current script could take a while to run.

## Results
### VBA Script
Using the starting script provided, below are the steps and resulting code created:
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
    For Row = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(Row, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(Row, 1).Value = tickers(tickerIndex) And Cells(Row - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(Row, 6).Value
                        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(Row, 1).Value = tickers(tickerIndex) And Cells(Row + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(Row, 6).Value
            
        End If

            '3d Increase the tickerIndex.
        If Cells(Row, 1).Value = tickers(tickerIndex) And Cells(Row + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next Row
    
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
```
### Results Analysis
The refactoring can be considered successful, as the time it has taken to run the script for 12 tickers has gone down from ~1.01 seconds to ~0.16 seconds.

![VBA_Challenge_2017_Old](https://user-images.githubusercontent.com/86527135/126080499-0033d76a-1b89-4f5a-88d0-1c087af56919.PNG)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/86527135/126080503-f4c89e48-96c1-4753-84cf-81633a1464f5.PNG)

![VBA_Challenge_2018_old](https://user-images.githubusercontent.com/86527135/126080506-843bf29b-9dca-4024-b89e-36681e2dbfa5.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/86527135/126080507-acf0b4ac-7508-4d9f-970c-e2e32b150d23.PNG)

## Summary
### What are the advantages or disadvantages of refactoring code?
Some of the advantages to refactoring code include a clearner script, which should make it easier for others to understand and use the code in the future. For this instance, the refactoring also made the script run faster, which will be necessary if we are to expand the data set from 12 tickers to hundreds. 

A disadvantage to refactoring code you're unfamiliar with could include altering sections that seem to be redundant, but had a purpose and therefore creating possible bugs in the future. This is why leaving notes and comments in the script is so important.

### How do these pros and cons apply to refactoring the original VBA script?
The script definitely runs faster, and will work for more tickers, making it an advantage overall. However, by looping through the code in one go and relying on the tickers to be grouped together and chronologically, we could run into issues if the data was sorted differently. So while the script does what we've asked it to do, further refactoring could be useful if we want to make sure it works with messier data sets.
