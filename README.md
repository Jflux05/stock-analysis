# stock-analysis
Analyzing Stocks with VBA

## Overview of Project

### Background & Purpose
Steve, our client, has requested some help to create an efficient process to analyze green energy stocks. Our client's goal is to utilize this analysis to help his parents diversify their investment portfolio and improve their investment strategy with data backed decision making. Currently, Steve's parents' entire holdings are comprised of a single stock, Daqo (DQ).

Our source data included ticker data from 12 different green energy stocks, including DQ, for the years 2017 and 2018. We utilized an easy to use click-button that would analyze the volume and returns of those stocks in a given year, and display them in an ease-to-read format. After a working subroutine was created, our goal was to optimize the efficiency of that macro by refactoring its code.

## Results
From our analysis we can see that the Steve's parents should diversify their portfolio beyond just DQ. We were able to see the below satisfactory performance from  DQ in both years. In comparison to the other 11 stocks, DQ is one of the lowest performers. Almost all of the stocks we tested had positive returns in 2017, and only two stocks showed positive performance in 2018: ENPH and Run (each showed over 81% positive returns). Overall, this small snapshot of the green energy sector shows that there is significant volitility in this sector of the market. Additionally, this should provide Steve (our client) with a visual representation to show and convince his parents to diversify their portfolio. 

### Original Code Results
![Original Code Results 2017 snapshot](https://github.com/Jflux05/stock-analysis/blob/6389538d0d05b58aefd698f4338ded3d135f680e/Resources/2017%20Original%20code%20results%20snapshot%20.png)
![Original Code Results 2018 snapshot](https://github.com/Jflux05/stock-analysis/blob/615380b6a6d02f069ddb5667ba178cd0c9b0ef62/Resources/2018%20Original%20code%20results%20snapshot.png)

![Original Code 2017 timing screenshot](https://github.com/Jflux05/stock-analysis/blob/cdd81bb602d73b0a99058417d806b0822f5f511c/Resources/Original%20code%202017%20timing%20screenshot.png)
![Original Code 2018 timing screenshot](https://github.com/Jflux05/stock-analysis/blob/cdd81bb602d73b0a99058417d806b0822f5f511c/Resources/Original%20code%202018%20timing%20screenshot.png)

### Orginal Code Used
![Original Code Snapshot](https://github.com/Jflux05/stock-analysis/blob/bbe4b71a669bd76c53a2459216b761b593d82b61/Resources/Original%20Code%20snapshot.png)

### Analysis of Original Code
By adding in code to test the runtime, we could see that it would be beneficial to refactor the code, so it could be more efficient when applying the code to larger datasets. The code is totally functional as is, however it took a very long time to run, over 72 thousand seconds for each year. At this rate, the code is not efficient and would cause a very long wait time for it to run if used at a larger scale. It should be refactored to run efficiently. 

### Refactored Code & Results
![Refactored Code screenshot part 1](https://github.com/Jflux05/stock-analysis/blob/5618e2daae9dea2231407fd59b0ef6b78bcfecc5/Resources/Refactored%20Code%20part%201.png)
![Refactored Code screenshot part 2](https://github.com/Jflux05/stock-analysis/blob/5618e2daae9dea2231407fd59b0ef6b78bcfecc5/Resources/Refactored%20Code%20part%202.png)

By refactoring the code, (introducing the "tickerIndex" variable along with 3 output variables to our subroutine) we are able to avoid nested "For Loops" in our code. Thus making it much more efficient (and ready for larger datasets), which can be seen in the timing snapshot below:

![refactored code timing snapshot 2017](https://github.com/Jflux05/stock-analysis/blob/5618e2daae9dea2231407fd59b0ef6b78bcfecc5/Resources/VBA_Challenge_2017.png)
![refactored code timing snapshot 2018](https://github.com/Jflux05/stock-analysis/blob/5618e2daae9dea2231407fd59b0ef6b78bcfecc5/Resources/VBA_Challenge_2018.png)

## Summary
It was clear that refactoring code within VBA subroutines can produce the desired results much quicker upon execution. However, it does run the risk of potentially breaking or duplicating existing code. Given the improvements in run time, the benefits of refactoring the code are self-evident, but at the same time, shouldn't be utilized if under a time crunch as there is a risk of running into a variety of time consuming obstacles and significant debugging. 

### Original Code vs. Refactored Code
The benefits of the running the original script are significantly outweighed by the benefits of the refactored script. The original script is SIGNIFICANTLY slower than the refactored code and will experience challenges if Steve wants to analysze larger dataset in the future. Even when adding in the formatting script to the refactored code, it is still significantly faster than the original subroutine and then having to run the formatting subroutine after each time. The refactored script is much more comprehensive and complete when compared to the original code used. 
