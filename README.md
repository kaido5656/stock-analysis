## Purpose
  The purpose of this assignment was to assist Steve with macros using VBA in Excel to analyze various stocks and their attributes. Though the initial version of this code worked fine with 12 stocks, though to apply real life applications and many other stocks the code might benifit from some refactoring. So for the purpose of this assignnment I will refactoring a sample code to see if it may run faster, more effieciently, or any other results.

## Results upon refactoring

  Some of the ways the existing code was refactored was by declaring arrays to hold the value of every ticker and its associated values outside of the for loop it was incremented in. In doing that the loop was freed from having to hold every value it has analyzed and able to focus on just the one tickerindex at a time. Another way the code was refactored for effieciency was by removing the nested loop that was used before in  the previous macro. This was done by initializing the tickerindex outside of the loop to hold a value of 0 and creating a for loop seperately from the main loop, to initialize totalvolumes() = 0. the rest of the code was largely left untouched.

### Code that had undergone refactoring
```
'1a) Create a ticker Index
    
    Tickerindex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For j = 0 To 11
    
        tickerVolumes(j) = 0
        
        
    Next j
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        
        '3a) Increase volume for current ticker
        
        tickerVolumes(Tickerindex) = tickerVolumes(Tickerindex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(Tickerindex) And Cells(i, 1).Value = tickers(Tickerindex) Then
                
                tickerStartingPrices(Tickerindex) = Cells(i, 6).Value
            
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i + 1, 1).Value <> tickers(Tickerindex) And Cells(i, 1).Value = tickers(Tickerindex) Then
            
                tickerEndingPrices(Tickerindex) = Cells(i, 6).Value
            
            
            '3d Increase the tickerIndex.
            
                Tickerindex = Tickerindex + 1
            
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All stock analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
```
## Impressions of refactoring

  Some of the benefits of refactoring code lead to code behaving more effiencently, such as compiling faster, and leading to less probable errors by segmenting code. Another benefit of this practice leads to understanding the code on a greater level, by being able to glance at code and understand which part does what as opposed to having the code intertwined and reliant on each other. However some of the drawbacks for refactoring are that they require a great understanding of the language and its nuances to be able to even see that the code could indeed be refactored. In this instance the code being refactored is not that large but I could see how a larger bit of code could take a lot of research and time.

## Results
  A few of the advantages of the newly refactored code was that it led to much faster compile times as opposed to the original code, also that each new code block was relatively seperate from one another and easy to grasp. This could lead to the code being troubleshooted easier or edited in the future for multiple stocks or other purposes with ease. The only drawback I encountered with refactoring the code was to truly understand the base code and how to go about editing it, the process took me quite to finally grasp. An advantage of the original code was that it felt more natural upon its creation, to say that it felt like a rough draft that could do the processes asked of it. A disandvatage was that it was not entirely to nice to look at afterwards and I can see where it would have benefited from another pair of eyes, as probably the code was not as polished and easy to understandat a glance. The run times of the code were also bit long to compile.

#### The refactored code run times
    
    
                
