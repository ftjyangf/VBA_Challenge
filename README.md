# **VBA_Challenge**

## the overrall of the Prject

### Steve wants to perform an analysis basic on the stock market data, we help him to analysis the total trading volumes and the return of stocks in a past year. Steve was very happy about the information we provided, and he wants to treat all the stock data in this method. We felt the speed of our code is incapable to automate large data in short time, so we refactoried the code, and it performed well.

## Analysis Result 
### for 12 stock we analyized, we found that in 2017, only one stock has negative return, which is TERP, all other stocks performed well, the total daily volumne has very weak relation with return, since the correlation is -0.121, and it's negative, the smaller total daily volume stock is perfered.In 2018 majority of the sample stocks has negative return. with a mean of -8.9% and median of -12.7%. the correlation between total daily volume and return is postive strong, meaning with higer total daily volume the return is higer.

## Improvement from refactoring 
### the runing time of refactored code is about 0.3 second for both year, comparing with the runing time of original code is about 3 seconds, which is the result when excel is runing behind. the average runing time goes to 11 seconds when the excel is operating on the frond, we can see the active cells are working very hard and changing all the time, which makes us believe this is caused by the fact that the computer memery will store the data in the excel cell first then change it everytime it runs, which significantly delay the process. After refactoring, the step in original is cut and it complete in much shorter time.

## Advantage and Disadvantage of refactoring
### the adavantage is obvious, the process makes programmer to solve the problem in different method or more efficient method which could be more precise and clean, the code also could be faster run by computer, such as in this case. the disadvantage might make us harder to find the error. for example in this case, i used doubt data type instead of long for total daily volume, which always gave me wrong result, in orginal code i found the result because the code run is slower.
