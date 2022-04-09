# Renewable Stock Analysis 2017 vs 2018
# Challenge Overview
Using the code we assembled as part of our week two excercise as a base, the goal of this challenge was to refactor the VBA code in order for excel to assist in analyzing the returns on 12 renewable energy stocks for the years 2017 and 2018. While the results remained consistent between our in class exercise and the weekly challenge, the implementation of the additional code resulted in a slightly slower code run time. 
# Results
Using the week two in-class code output as a jumping off point and then following the instructions to insert the refactored coderesulted in the following:
* 2017 & 2018 Original Run-Time: 
![alt text](https://github.com/AMDavitt/stock-analysis/blob/main/Original%202018%20Output.png)
* 2017 & 2018 Refactored Run-Time:
![alt text](https://github.com/AMDavitt/stock-analysis/blob/main/Refactored%202018%20Output.png)
# Refactored Code Below:
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
    
    '1a) Create a tickerIndex variable and set it equal to zero before iterating over all the rows
    tickerIndex = 0

    '1b) Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.

    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    '2b) Create a for loop that will loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
    '3a) Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
    '3b) Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
    '3c) Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

    '3d) Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    '4) Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet
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

# Summary
* Refactored code has the potential advantage of creating improvements by reducing the amount of code needed, improved documentation that can be more easily interpreted by project collaborators, and finally can provide performance enhancements by implementing operational efficiencies. Disadvantages of refactoring code include potentially introducing uneeded complexity, new bugs, and confusion if version control isn't maintained. 
* I was able to succesfully run the refactored VBA code for stock analysis, it seems the additions actually slowed my runs down (.03s vs .1s for 2018 stock analysis). Given that the output for both versions of the VBA code were the same, it seems that the new additions added some uneeded complexity and ultimately slowing down the code.
