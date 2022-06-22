# Stock-Analysis
Performing an Analysis on how 12 stocks performed over the course of 2017 and 2018
## Overview
This project creates a more easily readable analysis of specific stocks Total Daily Volume and Return for a specifically chosen year. By refactoring code that had already been developed for this purpose, I was able to reduce the amount of times required to loop through the data. This resulted in a much shorter runtime for the analysis to complete. This project not only demonstrates my ability to program in VBA, but also my ability to analyze code for potential weaknesses and ways to improve.
## Analysis & Challenges
Originally, the code I had developed to analyze these stocks would loop through the data for each stock I wanted to analyze. This, however, proves to be rather inefficient for the data set I was working with. Because the data was sorted by the respective Ticker of a stock then by the date, this system of recursive loops proved unnecessary. Instead, I could simply utilize arrays to store my results as I traverse through my data in a single swoop. These arrays were initialized as follows:

`Dim tickerVolumes(11) As Long`
 
 `Dim tickerStartingPrices(11) As Single`
 
 `Dim tickerEndingPrices(11) As Single`
 
"tickerVolumes" was set as a Long to store the rather large numerical value of the Total Daily Volume a particular stock had throughout the selected year. To start, each of the values in this array needed to be set to 0. "tickerStartingPrices" tracked the price each stock started at during the selected year. Finally, "tickerEndingPrices" tracked the price each stock ended at during the selected year.
 
 To keep track of each individual stock, the following index was created and initialized to the value of 0:

`Dim tickerIndex As Integer`

`tickerIndex = 0`
  
 Next, I could begin with the for loop that would run through the entirety of the data one time (From row 2 until the previuosly calculated final row). Within this loop, I calculated the Total Daily Volume for a specific stock by incrementing the individual values of tickerVolumes array by the volume of the row I was on. This is displayed in the following formula:
 
 `tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value`
 
Next, I would check to see if the value I was working in was the first of its kind. To do this, I ensured that the Ticker of the current row did not match the previous one and that this Ticker also matched the corresponding ticker in the "tickers" array. This ensured that the current Ticker was the first of its kind and therefore the earliest date and also that the Ticker was recording the correct value. If the ticker value of the current row passed both of these tests, then it would be recorded as the "tickerStartingPrice" for a given "tickerIndex."

`If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then`

  `tickerStartingPrices(tickerIndex) = Cells(i, 6).Value`

`End If`

Similarly, I also ran a check to see if the ticker I was working with was the last of its kind. To do this, I checked that the Ticker of the current row did not match the next one and that this Ticker also matched the corresponding ticker in the "tickers" array. If the ticker value passed both of these tests, then it would be recorded as the "tickerStartingPrice" for a given "tickerIndex." Additionally, the tickerIndex would be incremented so that the program would be prepared for the next ticker value.

`If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then`
  
  `tickerEndingPrices(tickerIndex) = Cells(i, 6).Value`  
  
  `tickerIndex = tickerIndex + 1`
            
`End If`

Finally, with the for loop completed and my data collected, I was prepared to process it and display it in the "All Stocks Analysis" worksheet. By running a for loop from 0 to 11, I would be able to complete this process one stock at a time. On each stage of the loop, I would activate the "All Stocks Analysis" worksheet, create a row for the Ticker I was working with, display its Total Daily Volume, and then display its Return for the year. This last column was calculated by dividing the "tickerEndingPrice" by the "tickerStartingPrice"  and then subtracting 1. 

`For i = 0 To 11`

  `Worksheets("All Stocks Analysis").Activate`

  `Cells(i + 4, 1).Value = tickers(i)`

  `Cells(i + 4, 2).Value = tickerVolumes(i)`

  `Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1`
        
`Next i`

## Results
After running the refactored macro, the user is able to type in which year they would like to analyze. After making their selection, the program then stores and analyzes all of the data for each individual Ticker. This analysis then gets displayed in a stylized table. Additionally, the recorded time to run this macro is displayed in a message box. While this is not particularly important to a typical use, a developer is able to use this information to see how thier changes have improved their program's overall efficiency. Below, I have displayed the resulting charts and runtimes of my new macro for both the 2017 and the 2018 spreadsheets.

2017 Stocks Analysis: Original
![2017 Stocks Analysis: Original](https://github.com/waciciarelli/Stock-Analysis/blob/main/Resources/2017%20Stock%20Analysis%20Original.png?raw=true)

2017 Stocks Analysis: Refactored
![2017 Stocks: Refactored](https://github.com/waciciarelli/Stock-Analysis/blob/main/Resources/2017%20Stock%20Analysis%20Refactored.png?raw=true)

2018 Stocks Analysis: Original
![2018 Stocks Analysis: Original](https://github.com/waciciarelli/Stock-Analysis/blob/main/Resources/2018%20Stock%20Analysis%20Original.png?raw=true)

2018 Stocks Analysis: Refactored
![2018 Stocks: Refactored](https://github.com/waciciarelli/Stock-Analysis/blob/main/Resources/2018%20Stock%20Analysis%20Refactored.png?raw=true)

As you can see, this macro is capable of creating a visually appealing and understandable chart the same as it was before while also cutting down on the runtime. In both the cases of the 2017 and 2018 spreadsheets, the runtime of the refactored macro was around 1/10th that of the original.

This code does, however, have some notable drawbacks to the original. While the code may be a bit shorter, it is not quite as simple to understand visually how this new code processes. This would only be further compounded if it was being looked at by a second developer or somebody without an extensive knowledge of the spreadsheets it works with. Additionaly, this refactored code only works when the spreadsheet is sorted by Ticker first, and then date second. If the Tickers were not sorted, the wrong values could be attributed to the Starting and Ending Prices. Furthermore, the program would run into a runtime error if it encountered a ticker again after it had already encountered it. The Starting and Ending Prices would be recorded multiple times and with different values, the total volume would not accumulate properly, and finally, the index would go beyond what the arrays had accounted for and would give a runtime error for the program.

## Summary
What are the advantages or disadvantages of refactoring code?
Refactoring code allows for a programmer to revisit old code and to make improvements as to its efficiency and effectiveness. Refactoring also gives developers deeper insight into how to approach different problems and also allows for them to continuously be improving both themselves and their code. However, refactoring isn't a perfect process. Refactoring is still subject to human limits and errors. Not every developer has the time to refactor their old code. Additionally, refactoring can lead to errors within programs when not all of the code is taken into consideration. It is quite easy to change one variable and have it lead to a runtime error, which would only take further time to hunt down and correct.

How do these pros and cons apply to refactoring the original VBA script?
In my case, I saw all of these pros and cons when refactoring my own VBA script. I was able to see how sorting data before running such programs can drastically cut down on necessary code and allow for much more efficient runtimes. I do feel that I have grown as a programmer from this experience. On the other hand, I was able to see how quickly changing small things from my old code can snowball into runtime errors. By forgetting to rename a variable here or there, I found myself having to hunt down exactly where the issues with my program were while sifting through all of the code.
