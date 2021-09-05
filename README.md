# An Analysis of Green Energy Stocks using Excel and VBA

## Overview of Project

### Purpose

This analysis is to help Steve, a recent graduate with a finance degree. His parents are passionate about green energy and have decided to invest all their money into DAQO New Energy Corp ($DQ) without doing much research. Steve has promised to look into DAQO stock for his parents but is concerned about diversifying their funds, so he wants to analyze a handful of green energy stocks in addition to DAQO's stock.

Steve has created an excel file containing the stock data and has asked us to help him analyze it. By using Excel and VBA, we automate the analysis using code and built-in macros to run scripts that finds the total daily volume and yearly return for each stock. The results will help Steve show his parents the performance of the green energy stocks and help them decide if DAQO is a good investment. 

In this anlysis, we will use the results of our script to compare the stock performance between 2017 and 2018. In additon, we have refactored the original code and will analyze and compare the new refactored code versus the old code, as well as discuss the difference in execution times between the two.

## Results

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

### Comparing the stock performance between 2017 and 2018

#### 2017

![VBA_Challenge_2017_results.png](https://github.com/alexhuynh0530/stock-analysis/blob/main/Resources/VBA_Challenge_2017_results.png)

#### 2018

![VBA_Challenge_2018_results.png](https://github.com/alexhuynh0530/stock-analysis/blob/main/Resources/VBA_Challenge_2018_results.png)

When comparing the performance of the green energy stocks shown above, we can see that 2017 had more **green** than 2018 (excuse the pun). You'll see that in 2017, DAQO outperformed the group of green energy stocks, returning 199.4% on the year. TERP was the only stock that had negative returns for 2017, although it was a loss of less than 10%.

In 2018, only 2 stocks had a green year, ENPH and RUN, both returning more than 81% while the rest of the group, including DAQO New Energy Corp ($DQ), had negative returns. In addition, DAQO had the worst returns out of the whole list of green energy stocks in 2018 returning -62.6%. 

To improve the analysis, please see the average return from 2017 and 2018 below.

#### Average Return from 2017 and 2018

![VBA_Challenge_avg_return.png](https://github.com/alexhuynh0530/stock-analysis/blob/main/Resources/VBA_Challenge_avg_return.png)

As you can see, the top 3 stocks that had the best returns was ENPH, SEDG, and DQ returning 105.7%, 88.4%, and 68.4% respectively. You can conclude that ENPH is the best stock to own from the data given. In addtion to having the best average returns from 2017 and 2018, ENPH also had consistent gains in both years returning 129.5% and 81.9%. 

When looking at DQ stock, although it made the top 3 stocks with highest returns, you'll see that it is highly volatile. It had the highest return in 2017 but the worst returns in 2018. This doesn't seem like an investment that Steve's parents should invest in especially at their age.

Please note, there are some limitations to this dataset as the data is only limited to 2017 and 2018. Global news, company specific news, and many other factors could have affected the performance of a stock in a particular year. Therefore, analyzing data with more years would improve the anlaysis.

### Comparing the stock performance between 2017 and 2018

### Summary

There are some limitations to this dataset that include:

- Data limited to only 2010-2017, fresher data could help with more recent years
- Lack of data about the form of marketing and promotion used in executing the campaign fundraising (i.e. email marketing, social media, etc.)

Another graph we could create could be another line graph of Theater Outcomes by Launch Date using percentages. As noted in the Analysis of Outcomes by Launch Date, the most successful campaigns launched in May. However, May also had the most failed campaigns by quantity. If we used percentages, we would see that May had the highest percentage of successful campaigns (67%) followed by June (65%), and December had the highest percentage failed (47%) followed by October (43%). To enhance this chart even further we could filter on goals with about $10,000 since Louise is budgeting over $10,000 for her campaign.
