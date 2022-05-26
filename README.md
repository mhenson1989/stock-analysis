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


![**2018 Stock Analysis Snapshot**](https://github.com/mhenson1989/stock-analysis/blob/main/Resources/AllStocks_2018.PNG)

![**2017 Stock Analysis Snapshot**](https://github.com/mhenson1989/stock-analysis/blob/main/Resources/AllStocks_2017.PNG)



### **Execution Time Analysis - *Reviewing Refactored Code***
Within my original code, my analysis time for both 2017 and 2018 averaged just under 2 seconds. Within my refactored code, I was able to get the code to run on average between 0.13 and 0.14 seconds for both the 2017 and 2018 stock analyses. This is significant, since I expected my run time to decrease by at least a factor of 12, however, the actual code runs slightly faster than my initial expectations.
Shown below, you can see the difference between the Original All Stocks Analysis Run Time and the Refactored All Stocks Analysis Run Time for 2018.  

![Original Run Time](https://github.com/mhenson1989/stock-analysis/blob/main/Resources/VBA_OriginalCode_2018.PNG)

![Refactored Run Time](https://github.com/mhenson1989/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

### **Challenges and Obstacles of Refactoring**
I was confronted with significant challenges with refactoring this code, particualry with the code block below. For the output of the Total Volume and Return line, I continued to get an error that stated: Compile Error. Expected Array. After reviewing, I realized that I needed to refine my tickerIndex variable and use that to replace my original nested loops. Once I was able to understand that the tickerIndex variable was essentially replacing my nested loops, thus allowing the code to run through only the 3,014 lines, versus 3,014 lines for each 12 tickers, the code refactoring made a lot more sense. 
![Refactored Code with Errors](https://github.com/mhenson1989/stock-analysis/blob/main/Resources/CodeSnippet_Step4.PNG)


### **A Final Comparison**


1. *What the advantages or disadvantages of refactoring code?*
	- When considering refactoring code, one advantage is to streamline the overall time spent analyzing the data. In this case, my run time was decreased by a quotient of 12 because I eliminated the nested loops and allowed the code to run through 3,014 lines versus approximately 36,000. I do think there's a cost time versus reward consideration when thinking about refactoring code. As an example, this refactoring process took me a full week to code and change my run time from a second to a little less than half a second. In a real world scenario, I might decide that the time spent refactoring is not worth the time saved. If the data set was significantly larger, or we were looking to scale the analysis, I could see this process as an advantage, however, for such a small data set, I think the time spent became a disadvantage.  


2. *How do these pros and cons apply to refactoring the original VBA script?*

	- **Pros:** Refactoring the original VBA script provides the opportunity for a cleaner script with less for loops. This creates a higher liklihood for outside readers to understand your script, or for the original coder to come back after a length of time and dive back into the logic seamlessly. It also has the liklihood of faster analysis run times with your data set, which can be advantageous for large data sets. 

	- **Cons:** The amount of time spent to refactor and trouble shoot the code was significant and does not outweigh the time shaved from the run time. 


