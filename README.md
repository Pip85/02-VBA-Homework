# **practice-vba**

Student project - Analyze stock price data using VBA.

## **software/tools used**

* Microsoft Excel<br>
* Visual Basic<br>

## **resources**
* Background and datasets provided as part of Georgia Tech Data Analytics Boot Camp:<br>
* Trilogy Education Services © 2020. All Rights Reserved.<br>
* Resources/alphabetical_testing.xlsx
* Resources/Multiple_year_stock_data.xlsx

## **project background**

* In this homework assignment you will use VBA scripting to analyze real stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the challenge tasks.

* Create a script that will loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* You should also have conditional formatting that will highlight positive change in green and negative change in red.


### **BONUS**

* Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
* Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.


## **acknowledgement**

* Background and datasets provided as part of Georgia Tech Data Analytics Boot Camp:<br>
* Trilogy Education Services © 2020. All Rights Reserved.<br>

* Project Author:  Valerie Pippenger - https://github.com/Pip85

## **process**

* VBA code runs through stock data for a worksheet each for 3 years (2014-2016). A summary is created for each worksheet showing Ticker, change in stock price for the year, percentage change during the year and sum of stock volume during each year. 
* Yearly change data is formatted in green for increases and red for decreases. On the right of the ticker summary for each worksheet, the code pulls the data for greatest percentage change increase, greatest percentage change decrease and greatest stock volume.

