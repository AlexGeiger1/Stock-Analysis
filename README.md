# Stock-Analysis

Stock analysis of VBA

# Overview 

The purpose of the Stock Analysis Challenge was to research performance on the stock market for 12 stocks for two past years with improved efficiencies by refactoring code that performed the same analysis.  Using VBA to aid with the analysis, the return percentages for each stock are calculated and daily trade volume tallied, allowing for comparisons between years. The analysis also relays the length of time taken to pull and summarize the data, which aids in comparing how long the code takes to run to try to optimize the performance when using the tool for large volumes of stock data.   

 

# Results 

## Returns 

As can be seen in the below tables for 2017 & 2018, the cells for results are color-coded percentage returns. Green cells are the stocks that gained value, and the red cells are stocks that lost value. This color-coding makes it easy to see that 2017 was primarily gains, while in 2018,  most lost value. Stocks with Ticker ‘RUN’ had the highest increase in returns from 2017 (5.5%) to 2018 (84%). ‘ENPH’ stock was the only other that had gains in both years.   

Using conditional formatting in VBA is what was used to color cells.  An IF-THEN statement highlighted the returns for either a negative or positive relationship to value using the ‘interior’ function. This can be seen in the following section of code:    

![Conditional Format](https://github.com/AlexGeiger1/Stock-Analysis/blob/main/Resources/firstscreenshot.png) 

 

## Run Time  

As can be seen in the pop-up boxes within the tables A & B below, in 2017, the original stock analysis code ran in .719 seconds while 2018 took .804 seconds.     The refactored code resulted in 2017 analysis time of .094 seconds and .086 seconds for 2018, a great improvement over the original code’s time (tables C & D below).   

 

Table A 

![PreVBA_Challenge_2017](https://github.com/AlexGeiger1/Stock-Analysis/blob/main/Resources/PreVBA_Challenge_2017.png) 

 

Table B 

![PreVBA_Challenge_2018](https://github.com/AlexGeiger1/Stock-Analysis/blob/main/Resources/PreVBA_Challenge_2018.png) 

 

Table C 

![VBA_Challenge_2017](https://github.com/AlexGeiger1/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png) 

 

Table D 

![VBA_Challenge_2018](https://github.com/AlexGeiger1/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png) 

 

  

# Summary 

The advantages of refactoring the code include increasing the efficiency code by reducing execution times as well as potentially simplifying or reducing lines of code. The disadvantages of refactoring  are potentially causing additional errors that were not present in the beginning of the code. This results in more glitches that may cause unintended changes and take some time to troubleshoot. 

In the current code, the refactoring helped reduce the time that was needed to execute the code and increased the efficiency by cycling through the data less times in comparison to the original code. This means that the code can run at much quicker times with greater data loads. The con of refactoring means that a lot of time was needed to rewrite and troubleshoot the newly refactored code to make sure it ran correctly.  
