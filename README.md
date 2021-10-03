# Stock-analysis
## Overview of the project
Inside the financial world, the stocks are the things that fluctuate the most. Most of the time stocks follow a pattern according to *sometimes* news of the world or inside the company. Knowing the historical behavior of a stock is really important. It helps us know the accurate moment *if you want* to buy or even sell a stock.  
After the DQ Analysis done for Steve to study his parent´s stock, he wanted to know the same information but for all the stocks.
## Purpose
Steve wanted us to analyze 12 different tickers in over 3000 rows in our Excel. The way we analyze it was the collect 2 main data to determine the best stocks with the best results comparing them to their volume.    
-We collected the total Volume for each stock summing up all of the daily volume.  
-And we also collected the return of gain or loss of each stock.  
By the end of the project Steve will show his parents which stocks performed the best in 2017 and 2018.  Ans since they don´t code we added to buttons to the Excel. One button to run the analysis and one to clear it. 
## Results
After working on the analysis we found out that the difference between years is big.  
### 2017
In 2017 most of the tickers or *stocks* had a positive return. Some stocks had a increased percentage of up to **199.4%**.  
There was only 1 stock (TERP) who had a negative percentage return.  
Here´s the data for 2017:  
![Data Analysis fro 2017](https://github.com/ManuelRuizF/stock-analysis/blob/main/resources/An%C3%A1lisis2017.PNG)

### 2018 
In 2018 most of the tickers had a negative return! The most negative was reached down to -62.6%. And only 2 stocks had a positive return: ENPH (81.9%) and RUN (84%). 
Here´s the data for 2018:  
![Data Analysis 2018](https://github.com/ManuelRuizF/stock-analysis/blob/main/resources/An%C3%A1lisis2018.PNG)

### Running times
We also coded so we can get back the running times of the original and th *refactored* (summary) code. In the following pictures we will compare the runnig times of both.  
#### 2017
In the following images we will see the difference between running times in  2017 between the original (1) and the refactored code (2).  
![Run Original code](https://github.com/ManuelRuizF/stock-analysis/blob/main/resources/tiempo2017original.PNG) ![Run ref code](https://github.com/ManuelRuizF/stock-analysis/blob/main/resources/tiempo2017refactored.PNG)
Refactored shortened the runnign time.
#### 2018
In the following images we will see the difference between running times in  2018 between the original (1) and the refactored code (2).
![Run original code](https://github.com/ManuelRuizF/stock-analysis/blob/main/resources/tiempo2018original.PNG) ![Run refac time](https://github.com/ManuelRuizF/stock-analysis/blob/main/resources/tiempo2018refactored.PNG)  

Refactored shortened the runnign time.
## Summary  
One extra challenge for this work was to refactor the code used in the prior work for Steve. To refactor a code means to change the *syntax* but without manipulating the instructions of it. The purpose of refactoring is to make the code easier to read, understand, and perform.  
Since we didnt want to copy and paste it several times we created 3 output arrays (Index, Volumes, prices). 
The best **advantage** of refactoring code is that it gives quality to it. It runs faster, and better. Also takes less space in the RAM memory. With a shorter code, it becomes easier to read and understand. In the VBA code we didnt repeat the code 12 times for each ticker, instead we array the 3 data we needed. It took less code and it ran way faster than the original.  
For the **disadvantages** when you refactor a code you have to tested a lot of times for every possible scenario. You got to go through everything that could fail. To refactor a code could take a lot of time and if your are paying for it it could cost you a lot of things don´t work out the first time.  
For our VBA code, we needed a different logic to do it. In the original one we used nested loops and in this one arrays, so it took us more time to develop the logic of the code.  


