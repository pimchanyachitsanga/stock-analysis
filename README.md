# Green Stock Analysis

## Overview of Project

### Background

A friend of mine, Steve has asked me to perform green stocks analysis (portfolio consists of 12 stocks) for his parents to find the stocks' annual return and daily volumn traded in 2017 and 2018. The tool I utilized for this analysis is the Visual Basic Appication in Excel. Upon the completion of the analysis, I was able to analyze the best investment option for Steve's parents.

### Purpose

The purpose of this project is to create a more efficient way to analyze at multiple stocks using VBA. After the initial analysis of green stocks , it is clear that there are more efficient ways to analyze the data by refactoring the initial code.

## Results

### Refactoring the Code

I have refactored the initial code to include a for loop through the data and collect all of the information all at once. In order to do this, I created a ticker Index and created three output arrays; tickerVolumes, tickerStartingPrices, and tickerEndingPrices. Then loop through the arrays to output the Ticker, Total Daily Volume, and Return by by using a variable called the tickerIndex.

### Refactored Code

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
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
        
        For i = 0 To 11
            
            tickerVolumes(i) = 0
        
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value            
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If
        
        '3c) Check if the current row is the last row with the selected ticker
        'If  Then
            
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
            '3d Increase the tickerIndex.
                 
                 tickerIndex = tickerIndex + 1
            
        End If
        
    Next i
        
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1

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

### Initial Code

Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
      startTime = Timer

    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
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

    Worksheets(yearValue).Activate

    'get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    Dim startingPrice As Single
    Dim endingPrice As Single

    For i = 0 To 11

        ticker = tickers(i)
        totalVolume = 0

        Worksheets(yearValue).Activate

        'loop over all the rows
        For j = 2 To RowCount

            If Cells(j, 1).Value = ticker Then

                'increase totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(j, 8).Value

            End If

            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                startingPrice = Cells(j, 6).Value

            End If

            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value

            End If

        Next j

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

Sub formatAllStocksAnalysisTable()

 'Formatting
     Worksheets("All Stocks Analysis").Activate
     
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
    If Cells(i, 3) > 0 Then

            'Change cell color to green
            Cells(i, 3).Interior.Color = vbGreen

    ElseIf Cells(i, 3) < 0 Then

            'Change cell color to red
            Cells(i, 3).Interior.Color = vbRed
            
    Else
    
            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

    End If

Next i
    
End Sub

## Summary

The analysis are completed much faster with the refactored code. Run times for each are shown below.

### Run-time for initial code in 2017 and 2018

![Initial Run-Time 2017](https://github.com/pimchanyachitsanga/stock-analysis/blob/master/Green_Stock_2017.PNG) 

![Initial Run-Time 2018](https://github.com/pimchanyachitsanga/stock-analysis/blob/master/Green_Stock_2018.PNG) 

### Run-time for refactored code in 2017 and 2018

![Refactored Run-Time 2017](https://github.com/pimchanyachitsanga/stock-analysis/blob/master/VBA_Challenge_2017.PNG) 

![Refactored Run-Time 2018](https://github.com/pimchanyachitsanga/stock-analysis/blob/master/VBA_Challenge_2018.PNG) 

### Advantages and disadvantages of refactoring code in general 

The advantages of refactoring code are making the code more efficient and also avoid longer run times which can causes the program to crash. The disadvantages of refactoring code are the additional time spent and potentially run into more problems if refactoring incorrectly. 

### Advantages and disadvantages of the original and refactored VBA script 

The advantages of refactoring code in VBA script are that you can put the new code side by side to find more efficient ways then reuse majority of the code to create a refactored code. The disadvantages of refactoring code in VBA script are that it requires strong understanding of loops and syntax used which could cause additional time to making the refactored code work than to use the original code.
