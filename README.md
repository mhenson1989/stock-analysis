# **Refactoring VBA** - *An Analysis of 2017 and 2018 Stocks*

## **Overview of Project**

### **Purpose**
Within this Module 2 Challenge, I performed an analysis which compared two data sets of stock from 2017 and 2018. The analysis created specific output that broke down Total Daily Volume and rate of return for each Ticker listed within the data set. 
The purpose of this analysis was to assess which Tickers would be a safe investment, based on their return rates, and to use this analysis to make recommendations on ticker options. I then refactored this code in order to reduce the run time for the analysis. The overall goal in refactoring the code was to create a more efficient, readable and logic driven code. This would be especially important if this analysis became the standard for future and/or additional stock analysis.  

## **Results**

### **Stock Performance Comparison**
At an aerial view, we can see that the tickers analyzed in 2017 overall yielded a higher return rate than the same tickers analyzed in 2018. One can assert that 2017 was an overall successful year for the majority of the ticker stocks analyzed, with the exception of the TERP ticker stock, which has seen a consistent decrease return rate over the two year period. 
Looking at the data from a year over year perspective, I would recommend investing in both the ENPH and RUN ticker stock. In 2017, the ENPH ticker saw a return rate of 129.5%. Comparitively, though not as high as 2017, the ENPH ticker stock saw a return rate of 81.9% in 2018, which is significantly higher than other ticker stocks. 
Similarly, the RUN ticker stock has seen incremental growth on return from 2017 to 2018. Though it's return rate in 2017 was minimal (5.5%), it saw a substantial increase in it's return rate, year over year, jumping to 84% in 2018. 
Both of these stocks show promise for future yield, especially when looking at 2018 return rates for the other analyzed ticker stocks - all of which had a negative rate of return. 


![**2017 Stock Analysis Snapshot**](https://github.com/mhenson1989/stock-analysis/blob/main/Resources/AllStocks_2017.PNG)


![**2018 Stock Analysis Snapshot**](https://github.com/mhenson1989/stock-analysis/blob/main/Resources/AllStocks_2018.PNG)

### **Execution Time Analysis - *Reviewing Refactored Code***
Within my original code, my analysis time for both 2017 and 2018 averaged just under 2 seconds. However, within my refactored code, my analysis time decreased for my 2017 Stock Analysis by approximately 0.4 seconds, but my analysis time for my 2018 Stock Analysis actually increased by approximately 0.4 seconds. See snapshot below for the increased run time for the 2018 Refactored Stock Analysis Code.
![Refactored 2018 Run Time](https://github.com/mhenson1989/stock-analysis/blob/main/Resources/VBA_Refactored_2018.PNG)

### **Challenges and Obstacles of Refactoring**
I was confronted with significant challenges with refactoring this code, particuarly in regards to the output loop, especially when calculating my return rate column - this line continued to prompt an error code. I was unable to find a solution for this error, and therefore opted to tag this line as a note in my code and exclude it from my output. Additionally, I had challenges when outputting the loop for Volume (Column B). In many cases, I recieved an error code that said "Compile error. Expected Array". I also sometimes was able to produce output, however, the output would be equal to zero(0) for every line. I believe that the code should be written like the snapshot below, which leads me to believe that there is a potential error in my defined arrays. Unlike in the original VBA script, which also had bugs when run originally, I was ultimately unable to find a concrete logic in the steps of the refactored code, which made it difficult to troubleshoot. 
![Refactored Code with Errors](https://github.com/mhenson1989/stock-analysis/blob/main/Resources/CodeSnippet_Step4.PNG)


### **A Final Comparison**


1. *What the advantages or disadvantages of refactoring code?*
	- Though refactoring the stock analysis code shaved time from my macro, I would argue that ultimately the time spent refactoring the code, did not ultimately save time on the project overall. I could see how in a significantly larger script, it would be advantageous to refactor the code to create a more streamlined, efficient and readable script, the original script for our All Stocks Analysis was not overly complicated. When looking at refactoring it for more efficient code, I found that the steps were largely a duplication of the original code with a reliance on the tickerIndex. Though the tickerIndex was intended to simplify the for loops and conditionals, I found the nested loops to be overall, more difficult to comprehend than the original code. 


2. *How do these pros and cons apply to refactoring the original VBA script?*

	- **Pros:** Refactoring the original VBA script provides the opportunity for a cleaner script with less for loops.

	- **Cons:** Though it might have eliminated nested loops, the inclusion of the tickerIndex variable to streamline the four arrays, essentially creates a nested logic that is more difficult to understand - thus significantly increasing the time spent to refactor the code. 


