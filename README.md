# stock-analysis

## Overview:

The purpose of this project was to help Steve find the total daily volume and yearly return for 
12 energy stocks in 2018 and 2017. This was done by refactoring code in order to make it run faster. 
The reason behind this is because the old code was not as effecient as it could be and refactoring it 
could make it more effiencent. Doing this makes the code take less memory and improves its logic.


## Results:

The new refactored code executes steves goal more effeciently, which can be seen by comparing
the time it took to run the old and new code. 
![Original_Analysis_2017](https://user-images.githubusercontent.com/107213807/174387543-a8404e95-c679-4098-8aa3-5f4c0011ffe1.png)
![Original_Analysis_2018](https://user-images.githubusercontent.com/107213807/174387586-2b143566-b7d2-41a9-911a-1fc82a8db680.png)
The orignal times for 2017 and 2018 are noticibly slower than the the new code.
![VBA_Challenge_2017](https://user-images.githubusercontent.com/107213807/174387673-685897f3-f50a-4100-9efc-6fe996e9db9b.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/107213807/174387685-e5c76cef-6bf9-4f1a-b2bc-418e9fa47744.png)
The key difference that made the code run faster was the inclusion of a tickerIndex variable. Using this tickerIndex variable,
it was possible to create output arrays for the tickerStartingPrices, tickerEndingPrices, and tickerVolumes. The tickerIndex
variable was used for all three arrays in a for loop that finds the total daily volume for each index, as well as the starting
and ending prices for each index in order to calculate the yearly return.

## Summary:
