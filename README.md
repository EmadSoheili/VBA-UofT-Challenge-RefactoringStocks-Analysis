# VBA of Wall Street

***University of Toronto - Data Analytics Boot Camp - Module 2 - VBA of Wall Street***

---

## Overview of Project

This analysis is carried out in order to provide us with some insights about some stocks. Therefore with the help of these insights, we might be able to make a data-driven decision and choose the best one out of all. The required insights are:
  - **Total Volume** of each stock's shares traded which shows the level of trading
  - **Yearly Return"** of each stock that illustrate the level of profit or loss

*The VBA code also is **Refactored** in orderto become more efficient.*

---

## Results

### Stock Analysis
   
#### ***2017***

  - It seems like a great day for the stocks market. Almost all stocks price grew over the year, except one.
  - The market average growth among targeted stocks were more than 67%, which is very interesting.
  - First 2017 was a perfect time to enter the market.
  
  ![](/Resources/2017-StocksStatus.PNG) 
  
#### ***2018***

  - Despite year 2017, 2018 seems not a really great time to enter the market.
  - Most stocks experienced a major loss during 2018
  
  ![](/Resources/2018-StocksStatus.PNG) 
  
#### Comparison

  - Obviously 2017 was a fruitfull year, and 2018 was a year full of losses among the targeted stocks.
  - Total volume of trades in 2018 is a little more than 2017

#### Best Stocks to invest in

  - Since we need to choose best options among these 12 stocks, we need to point out the stocks that experienced two positive years in row.
  - Out best options are ***ENPH*** and ***RUN***
  
  ![](/Resources/2017-StocksStatus.PNG)   ![](/Resources/2018-StocksStatus.PNG) 
  
#### Time-Efficiency Comparison

We have refacored the code we wrote in the first place. This was achievable since a nested for loop (contains another for loop) has changed to a unnested for loops. This change is beneficial by decreasing the running speed of our Macros.

As illustrated below, the execution time has decreaced by nearly **83%**.

![](/Resources/2017-BeforeRefactoring.PNG)![](/Resources/VBA_Challenge_2017.PNG)

![](/Resources/2018-BeforeRefactoring.PNG)![](/Resources/VBA_Challenge_2018.PNG)

***This shows that the Refactoring was truly resultful.***
