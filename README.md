# stock-analysis

## Overview:

The purpose of this project was to help Steve find the total daily volume and yearly return for 
12 energy stocks in 2018 and 2017. This was done by refactoring code in order to make it run faster. 
The reason behind this is because the old code was not as efficient as it could be and refactoring it 
could make it more efficient. Doing this makes the code take less memory and improves its logic.


## Results:

The new refactored code executes steves goal more efficiently, which can be seen by comparing
the time it took to run the old and new code.

![Original_Analysis_2017](https://user-images.githubusercontent.com/107213807/174387543-a8404e95-c679-4098-8aa3-5f4c0011ffe1.png)
![Original_Analysis_2018](https://user-images.githubusercontent.com/107213807/174387586-2b143566-b7d2-41a9-911a-1fc82a8db680.png)

The original times for 2017 and 2018 are noticeably slower than the new code.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/107213807/174387673-685897f3-f50a-4100-9efc-6fe996e9db9b.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/107213807/174387685-e5c76cef-6bf9-4f1a-b2bc-418e9fa47744.png)

The key difference that made the code run faster was the inclusion of a tickerIndex variable. Using this tickerIndex variable,
it was possible to create output arrays for the tickerStartingPrices, tickerEndingPrices, and tickerVolumes. In order for the refactored
code to work, the output array tickerVolumes needs all of its indexes set to zero. This was done by using a for loop.

![tickerIndex_OutputArrays](https://user-images.githubusercontent.com/107213807/174390631-fcaa5032-ea72-4bb2-b2f0-7eb9ad26a630.png)

The tickerIndex variable was used for all three arrays in a for loop that finds the total daily volume for each index, as well as the starting
and ending prices for each index in order to calculate the yearly return. The code before the refactor did not have the tickerIndex 
variable, and used a nested for loop instead. Since the new code has the tickerIndex variable, there is no need for a nested loop
as the loop in the refactored code increases the tickerIndex by 1 after getting the needed cell values for each ticker. Inside the for loop, two if statements were used to get the correct
values for the tickerStartingPrices and tickerEndingPrices arrays.

![OutputArrays_Loop](https://user-images.githubusercontent.com/107213807/174390696-7214d0ea-9d11-42f4-bbd9-e7c2ff45aebc.png)

This makes the code more efficient as it only needs to run through a loop rather than a nested loop. The rest of the code stayed
pretty much the same as no changes needed to be made to the formatting script or the loop that puts the new values into the table.

## Summary:

Refactoring code has advantages and disadvantages. An advantage is that it there is a clear objective of the code already, which
cuts out some of the early process of writing code or starting from scratch. Refactoring code saves time as the necessary scripts
for the code to work may already be created. While refactoring can save a lot of time, it can also be difficult to understand the context of certain code depending on the
code being refactored. There are many elements that may lead to confusion when trying to refactor some else's code, as there might
not be comments explaining what certain code does. Overall, refactoring code is important as it not only improves code, but also is more efficient as you don't have to start from scratch.

The refactored and original VBA scripts have their own advantages and disadvantages. The refactored script is more efficient
as it uses a variable for the ticker indexes which is an advantage. The original code depends on a nested loop, which is
a disadvantage. The reason I consider this a disadvantage is the lack of flexibility. If more stock tickers were to be added or removed,
more lines of code would have to be edited compared to the new refactored code.


