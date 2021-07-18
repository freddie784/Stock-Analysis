# VBA Stocks Project

## Overview of Project and purpose

We were asked to create a macro that would analyze a dataset of stocks for clean energy companies. We wanted to condense the data to summarize the returns for each ticker that was given in the dataset. We specifically were looking at *Total Daily Volume and Yearly Returns*. For the Challenge we needed to refactor our code because the original did the job but in an inefficient manner.

## Results

When I ran the Macro to analyze the stocks, it was obvious that 2017 was a much better year to invest in stocks than 2018. The only two stocks that had consistent gains through both years were **ENPH** and **RUN**. As for the code, when it was first created took around .7 seconds to analyze a dataset. ***insert Picture*** This code was inefficient because there were three *for* loops each analyzing a different column on the **SAME** worksheet. Each column has 3012 rows of data which meant that after all three loops were done 9036 rows were analyzed. In the refactored code we created **one** loop that checked all **three** columns for the data we were looking for; therefore, cutting the amount of rows to loop through by *66%*. This refactor cut the code's time to execute by about 6 seconds ***See Below*** ***Insert Pic***

## Summary

### Advantages of refactoring code

I learned that refactoring code allows you to shorten the time it takes your computer to compute the results you need. It is a way to make your code more efficient therefore allowing the code's user to gain productivity.

### Disadvantages of refactoring code

A disadvantage to refactoring code is that you could be moving on to the next project or new problem you want to solve. Also if you are a company depending on how long it takes to find a solution to create a refactored code the cost of doing so may be more than it was worth.

### How these pros and cons applied when I was refactoring

When I ran the first code we made, it made me question my cpu's processing power cause it was taking so long. After I made the refactored code I realized how inefficient the original was. While refactoring the code I noticed that even though the problem of figuring out how to analyze the stocks was solved, it was a whole different project when it came to refactoring it. In this case though it was definitely worth it to refactor because it was not too difficult of a fix and it cut our time by *85.7%*!
