# Written Analysis: Stock Analysis with VBA
> Analyzing the stock market over few years and re-factoring the VBA solution code.

## Table of Contents
* [Overview of the project](#overview-of-the-project)
* [Purpose](#purpose)
* [Results](#results)
* [Analysis of stocks](#analysis-of-stocks)
* [Code comparison and it's output](#code-comparison-and-it's-output)
* [Execution time for the code](#execution-time-for-the-code)
* [Summary](#summary)


## Overview of the project
This challenge is to expand the dataset to include the entire stock market, over 2017-2018.
### Purpose
The purpose of this challenge was to help our client Steve, to analyse the green energy stock market for his parents. His parents are interested in the DAQO stocks and before investing into these, Steve wanted to run some analysis and check the performance of the DAQO stocks over the years in comparison to other green stocks. We helped Steve perform this analysis using VBA code to come up with the total daily volume and yearly return for each stock. In this challenge, we are re-factoring the same VBA solution code to loop only once and determine if the VBA script ran faster than the original, due to the code's increased efficiency.


## Results
### Analysis of stocks
The images below show the comparison of a dozen green energy stocks. The values ofcomparion include three groups of data:
*Ticker name
*Total Daily Volume for a given year
*Percentage of the yearly return for each stock in the given year


![Stock analysis 2017](./img/img_1.png)  ![Stock analysis 2018](./img/img_2.png)


As you can see,the stocks in 2017 had a high ratio of positive returns whereas 2018 returns show a completely opposite picture.Majority of the stocks had a significant drop in its returns. The DQ stock had a return value of nearly 200% in 2017 as compared to the negative 62% in 2018. 
These results indicate that the DQ stock trend is not stable and might be a risky investment for Steve's parents, based on it's yearly return.

If we now compare the results of the daily volume, it shows that DQ stocks had a low volume and high return in 2017. However, in 2018 the situation of the DQ stocks had reversed completely. Despite the trading volume being higher in 2018, the returns were very low and infact negative.
These results yet again confirm a risky investment in the DQ stocks for Steve's parents.

### Code comparison and it's output
The orignal VBA code included two loops. The code in a nested loop is switching back and forth between worksheets.

![Original code](./img/img_3.png)


The re-factored VBA solution code was consolidated into one loop. The code stays in the same loop, gathers and stores all the data in arrays.

![Refactored code](./img/img_4.png)


Both the scripts "AllStockAnalysis" and "AllStockAnalysisRefactored" have the same output in terms of it's values for each green stock.

### Execution time for the code

When the refactored code was executed against 2017 stock market data set, it ran in 0.109 seconds as compared to the original code that ran in 0.597 seconds for the year 2017, which is almost 6 times slower than the refactored code.



![Refactored code run time](./img/img_5.png)  ![Original code run time](./img/img_7.png)




When the refactored code was executed against 2018 stock market data set, it ran in 0.171 seconds as compared to the original code that ran in 0.605 seconds for the year 2017, which is nearly 4 times slower than the refactored code.



![Refactored code run time](./img/img_6.png)  ![Original code run time](./img/img_8.png)


I also noticed that for the same script,the execution time varied every time the code was run,however, overall the refactored script still ran faster than the original one.


## Summary
**1. What are the advantages or disadvantages of refactoring code?**<br />
* Our goal with refactoring is to use an existing code and restructuring it to improve it's overall efficacy and readablity , reduce execution time and reduce room for errors while still preserving it's functionality. It is cleaner and well organized, ulimately allowing us to reap the benefits as mentioned. <br />
* The downside however, is that it is very time consuming to understand which part of the code requires refactoring, whether new variables are required and at what step, whether the functionality can still be maintained and at the end to determine whether after refactoring the script, would the new code even run faster than it already was.Overall, it could endup being less efficient, time consuming and a challenging process.

**2. How do these pros and cons apply to refactoring the original VBA script?**<br />
* Clearly the refactored VBA code was faster in it's execution time due to some factors such as condensed loop that potentially decreased the processing memory required to process the data set. The whole process was pretty confusing and complicated,that led to spending a lot of time in testing the codes. There was a lot of debugging required as well at each step. VBA did not seem tobe very clear when communicating the errors, which caused a lot of waste of time during the refactoring process. The original code was still effective for the smaller dataset and it did execute the deseired output in a decent run time. However, it would have been challenging and less efficient to use the same code for a large dataset, considering it having a loop repeat multiple times through the large set of data. Overall,refactoring did help and was rewarding in its outcome at the end.

