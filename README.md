# Stock-analysis
# "Green_stocks Analysis."

## Overview of Project

### Purpose
 
The purpose of this project is use VBA applications to refactor a finance analysis, Green_stocks.
The main goal is loop through the stocks making more efficient, easier to comprehend and most important, run faster. 

## Analysis 

The first step of the refactoring was create three new arrays: tickerVolumes, tickerStartingPrices and tickerEndingPrices, using a tickerIndex equal to zero in order to access the current Index across the arrays I would be using.

Next, I used for loops to initialize a tickerVolumes to zero and loop over through the spreadsheet rows. Then, ‘If statements’ were used to define the starting and Ending Price finding the values. Three columns were added to the spreadsheet showing the ticker, total daily volume and returns after loop through the arrays.

![VBA_Challenge_1resouces](https://github.com/jeperes/Stock-analysis/blob/main/Resources/VBA_Challenge_1resouces.jpg)

In order to organize the table, the formatting syntax was used, changing colors, numeric formatting and autofitting the values.
 
![VBA_Challenge_2resouces](https://github.com/jeperes/Stock-analysis/blob/main/Resources/VBA_Challenge_2resouces.jpg)

## Results

The 2017 output show a promising and safe investments with positive returns. The investments Steve’s clients are most interest, DQ, have a high return percentage of 199.4%. However, in 2018, the investments are not successful compared to the previous year, with -62.6%. The images below shows the table with all results.

![VBA_Challenge_2017](https://github.com/jeperes/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.jpg)


![VBA_Challenge_2018](https://github.com/jeperes/Stock-analysis/blob/main/Resources/VBA_Challenge_2018.jpg)

# Summary

- What are the advantages or disadvantages of refactoring code?

Advantages

A big benefit of refactoring is definitely being able to improve the programming without adding new functionalities, but collecting the same information. Also, using less memory and running the macros really quick.

Disadvantages

The disadvantages of refactoring are the process can take a lot of time to complete and even correct an error. And working with an existent data is risk because you may have to change some functions that was using in the original script.

- How do these pros and cons apply to refactoring the original VBA script?

One advantage of refactoring was making the code more efficient and readable, but using the same functionalities used before, especially because of the comments explaining step by step. Also, the code run macro time decreased, what is great for time consuming.

The time taken to refactoring the programming was really long, making the process a little bit frustrating and not very efficient. Also, I faced some complex coding and took a quiet time to finally figure out the problem.  


