# stock_analysis 

## Objective/Purpose
The purpose of this assignment was to see whether refactoring VBA code can help with decreasing subroutine runtime and whether it was worth it. As part of classwork, a VBA subroutine was created to analyze stock data and highlight total daily volumes and the yearly return for each stock in the data. Building on this, the project was undertaken to refactor that code and see which one was better (comparing different aspects such as run time and readability).

## Analysis and Challenges

The VBA_Challenge.vbs file (part of the repository) was provided to begin with. This file consisted of sections of code that were to stay the same between older version of the subroutine and the refactored one. It also had some comments on the next steps.

The first thing was to create tickerIndex and set it to 0. A major part of refactoring was to eliminate the nested loops and use the creation of the following three output arrays to help with the final output. One of the output arrays was set as ‘Long’ while the other two were set as ‘Single’, seen in the following image:

<img width="472" alt="Code(Part1)" src="https://user-images.githubusercontent.com/92544151/149593498-44ee2a52-dc51-4464-8893-478fe00b616d.png">

A loop was then created for 2a to make sure that ‘tickerVolumes’ for each loop start at 0. To understand the refactored code for 2b, it should be mentioned that the following bit of code was taken from stack exchange to help get the number of rows to loop over:

<img width="364" alt="Total_Rows_Code" src="https://user-images.githubusercontent.com/92544151/149593530-1332345e-071a-449f-8033-817f68d21a3f.png">

From, 2b to 3d, it’s just one major for loop, iterating through the whole dataset and determining where one ticker ends and the other begins. If the previous row value did not have the same ticker name but the current one (i) did, then the values could be used for ‘tickerStartingPrice’. Along the same vein, if the current row value had the ticker name and the next one (i + 1) did not, then we could use it for ‘tickerEndingPrices’. The latter was also used to determine when the ‘tickerIndex’ value would increase. 

<img width="659" alt="Code(Parts2 3)" src="https://user-images.githubusercontent.com/92544151/149593549-454ce80a-e738-4171-a634-3a758d4ce5e2.png">

For step 4 of the analysis, it was simply another for loop, used to iterate through the arrays and outputting the values for Ticker, Total Daily Volume and Return. 

<img width="548" alt="Code(Part4)" src="https://user-images.githubusercontent.com/92544151/149593569-91a81065-3ab6-4657-8c4b-f51bf091a0e8.png">

The formatting was part of VBA_Challenge.vbs file provided as it stayed the same and had been written earlier as part of the classwork.

## Results

The following images show the pre-refactoring runtime for the subroutine:

<img width="257" alt="Pre-refactoring(2017)" src="https://user-images.githubusercontent.com/92544151/149593618-6e160617-6752-494d-bba9-45b19a9601e0.png">
<img width="258" alt="Pre-refactoring(2018)" src="https://user-images.githubusercontent.com/92544151/149593629-951f1dec-bf37-418e-985c-31d5e4465a26.png">

Compare that to the runtime after refactoring code:

<img width="252" alt="Refactored(2017)" src="https://user-images.githubusercontent.com/92544151/149593649-fb3e3915-37ef-44f9-ac81-81604c96a453.png">
<img width="245" alt="Refactored(2018)" src="https://user-images.githubusercontent.com/92544151/149593662-dfc3b2e3-a02c-4506-b850-1c4de1fc03b9.png">

The difference was more significant than expected. The refactored code was 5.56 times faster for the year 2017 and 5.52 times faster for 2018. 

Comparing the ticker performances between the years, the it is clear to see the stark difference between the amount of red and green in both. In 2017, all the tickers being analyzed had positive yearly returns, except TERP. DQ, ENPH, FSLR and SEDG had yearly returns greater than 100%. On the other hand, in 2018, only ENPH and RUN had positive returns. ENPH being the only stock to have performed well across both years. 

<img width="223" alt="Final_Analysis(2017)" src="https://user-images.githubusercontent.com/92544151/149593718-7e08aa43-2617-4e39-817d-14f49d62181c.png">
<img width="225" alt="Final_Analysis(2018)" src="https://user-images.githubusercontent.com/92544151/149593724-b7ec2146-ab6c-46fe-af48-30b93ee9da48.png">

## Summary

Looking at the results, it is easy to summarize that refactoring code led to much better runtimes for the subroutine. This shows that if minimizing the use of processing power is the goal, then refactoring code, as was done in this project, is definitely the way to go.  On the other hand, in my opinion, the refactored code will prove difficult to read and understand for others. 

The different in length between the earlier version of the subroutine and the refactored one was not significant. In the end, it all comes down to preference. Are you willing to sacrifice processing power to make the VBA code readable? Lean code or readable code? It should be mentioned that in some cases, lean code and readable code go hand in hand and there is no tradeoff but in other cases, like this project, it comes down to choosing one or the other. I chose to eliminate nested for loops and use extra output arrays and main loops with conditionals to make the subroutine more than 5 times faster. 
