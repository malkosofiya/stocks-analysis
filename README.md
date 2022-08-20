# VBA of Wall Street

## Overview of Project

The goal of the project was to find and analyze the total daily volume and yearly return for different stocks using VBA. The completed analysis helps decide which stocks are best to invest in and which to pass on. Additionally, the original script has been updated in order to make the report run faster and more efficient. 

## Code Overview

Although the original code worked well, the updated code is able to perform the same tasks faster for a larger set of data. By improving the logic of the code, we were able to reduce processing time from 0.25 seconds to 0.08 seconds. 

<img width="256" alt="Screen Shot 2022-08-20 at 5 09 39 PM" src="https://user-images.githubusercontent.com/110862261/185767754-e899f0cc-fec5-4974-912f-ddd3157ba8b9.png">

<img width="256" alt="Screen Shot 2022-08-20 at 5 13 50 PM" src="https://user-images.githubusercontent.com/110862261/185767805-b9fec65a-e5c0-43d5-b8b8-e99083749c29.png">

## Analysis and Results

At first glance, we can see that the majority of stocks performed better in 2017 versus 2018, with TERP being the only stock producing negative returns both years. In 2017 SPWR had the highest total daily volume with nearly 800 million of shares traded and an adequate return of 23%. In contrast, DQ had only 35 million shares traded but was able to produce almost 200% in return.

<img width="223" alt="Screen Shot 2022-08-20 at 4 50 27 PM" src="https://user-images.githubusercontent.com/110862261/185767201-f51d0d39-3fc2-4cd2-a2a3-ef6a761a041e.png">

In 2018 the only two stocks that had a positive return were ENPH and RUN. The previous year's high performers such as DQ and SEDG both turn out negative returns with -63% and -45% respectively.

<img width="225" alt="Screen Shot 2022-08-20 at 4 37 25 PM" src="https://user-images.githubusercontent.com/110862261/185767442-032adacc-baeb-4e5b-a4b2-d33520c94368.png">

When looking at the data from both years, ENPH might be a good option for investments since it performed consistently well both years with high returns and high number of daily trades.

## Advantages and Disadvantages of Refactoring Code

Refactoring helps to improve quality of the code therefore has some great advantages: 
* It makes code easier to understand
* It helps program to run faster and smoother
* It helps find issues such as duplicate code
* It can help fix bugs
That being said refactoring also has some drawbacks that should be considered: 
* It can be expensive 
* It can be time consuming

## How do these pros and cons apply to refactoring the original VBA script?

When working at this project, refactoring helped improve the runtime and efficiency of the code. However, refactoring code from its original script also took a lot of time and additional testing.
