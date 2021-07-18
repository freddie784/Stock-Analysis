# VBA Stocks Project

## Overview of Project and purpose

We were asked to create a macro that would analyze a dataset of stocks for clean energy companies. We wanted to condense the data to summarize the returns for each ticker that was given in the dataset. We specifically were looking at *Total Daily Volume and Yearly Returns*. For the Challenge we needed to refactor our code because the original did the job but in an inefficient manner.

## Results

When I ran the Macro to analyze the stocks, it was obvious that 2017 was a much better year to invest in stocks than 2018. The only two stocks that had consistent gains through both years were **ENPH** and **RUN**. 

As for the code, when it was first created took around .7 seconds to analyze a dataset.

![Screen Shot 2021-07-17 at 6 41 53 PM](https://user-images.githubusercontent.com/83510059/126054008-a8197502-8dd9-4206-941a-855f19ff7ce9.png)
![Screen Shot 2021-07-17 at 6 42 17 PM](https://user-images.githubusercontent.com/83510059/126054011-ce68d293-ca9d-4d18-aea8-c9d15f60be68.png)


This code was inefficient because it would check every row before moving on to the next ticker and adding the data to our results page. Each column has 3012 rows of data which meant that after the loop had finished running a total of 36,144 rows were analyzed. In the refactored code the loop went through the rows once for the data we were looking for; therefore, cutting the amount of rows to loop through by *92%*. This refactor cut the code's time to execute by about 6 seconds ***See Below***


![VBA_Challenge_2017](https://user-images.githubusercontent.com/83510059/126054019-a5384221-9733-4a16-aa2e-687eb9e31840.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/83510059/126054028-28b664e6-3694-4af9-8156-374ed8eca980.png)

The first code contained a nested *for* that counted the tickers equal to t 
>       For t = 0 To 11
>       
>           ticker = tickers(t)

so all the rows had to be checked before moving on to the next ticker. However, in the refactored code we made a tickerindex and equaled it to 0 ***"ticker index = 0"*** after that in the code it was checked to see if the it was the last instance of the ticker and if it was then we would move on to the next ticker without needing to go back to the first row and instead simply moving on to the next row.

>       If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then 'enter endingprice
>            tickerEndingPrices(tickerindex) = Cells(i, 6).Value
>           'Increase the tickerIndex.
>          tickerindex = tickerindex + 1
>       End if
>       next i

## Summary

### Advantages of refactoring code

I learned that refactoring code allows you to shorten the time it takes your computer to compute the results you need. It is a way to make your code more efficient therefore allowing the code's user to gain productivity.

### Disadvantages of refactoring code

A disadvantage to refactoring code is that you could be moving on to the next project or new problem you want to solve. Also if you are a company depending on how long it takes to find a solution to create a refactored code the cost of doing so may be more than it was worth.

### How these pros and cons applied when I was refactoring

When I ran the first code we made, it made me question my cpu's processing power cause it was taking so long. After I made the refactored code I realized how inefficient the original was. While refactoring the code I noticed that even though the problem of figuring out how to analyze the stocks was solved, it was a whole different project when it came to refactoring it. In this case though it was definitely worth it to refactor because it was not too difficult of a fix and it cut our time by *85.7%*!