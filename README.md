# An Analysis of Green Energy Stocks using Excel and VBA

## Overview of Project

### Purpose

This analysis is to help Steve, a recent graduate with a finance degree. His parents are passionate about green energy and have decided to invest all their money into DAQO New Energy Corp ($DQ) without doing much research. Steve has promised to look into DAQO stock for his parents but is concerned about diversifying their funds, so he wants to analyze a handful of green energy stocks in addition to DAQO's stock.

Steve has created an excel file containing the stock data and has asked us to help him analyze it. By using Excel and VBA, we automate the analysis using code and built in macros to run scripts that finds the total daily volume and yearly return for each stock. The results will help Steve show his parents the performance of the green energy stocks and help them decide if DAQO is a good investment. 

Furthermore, we have refactored the code and will analyze and compare the new refactored code versus the old code, as well as discuss the difference in execution times between the two.

## Results

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

### Comparing the stock performance between 2017 and 2018

![Theater_Outcomes_vs_Launch.png](https://github.com/alexhuynh0530/kickstarter-analysis/blob/main/resources/Theater_Outcomes_vs_Launch.png)

When comparing the performance of the green energy stocks given, we saw more **green** in 2017 versus 2018 (excuse the pun). You'll see that in 2017, DAQO outperformed the group of green energy stocks, returning 199.4% on the year. TERP was the only stock that had negative returns for 2017, although it was a loss of less than 10%.

In 2018, only 2 stocks had a green year, ENPH and RUN, both returning more than 81% while the rest of the group, including DAQO New Energy Corp ($DQ), had negative returns. In addition, DAQO had the worst returns out of the whole list of green energy stocks in 2018 returning -62.6%. 

### Conclusions made about the Analysis of Outcomes Based on Launch Date

- Campaigns launched in May had the highest number of successful campaigns
- Campaigns launched in December had the lowest number of successful campaigns

### Conclusions made about the Analysis of Outcomes Based on Goals

- Fundraising goals of less than $5,000 had the highest success rates
- We could also conclude that other factors other than fundraising goals play a part in a successful campaign as we saw goals set between $35,000 and $44,999 also having high success rates

### Summary

There are some limitations to this dataset that include:

- Data limited to only 2010-2017, fresher data could help with more recent years
- Lack of data about the form of marketing and promotion used in executing the campaign fundraising (i.e. email marketing, social media, etc.)

Another graph we could create could be another line graph of Theater Outcomes by Launch Date using percentages. As noted in the Analysis of Outcomes by Launch Date, the most successful campaigns launched in May. However, May also had the most failed campaigns by quantity. If we used percentages, we would see that May had the highest percentage of successful campaigns (67%) followed by June (65%), and December had the highest percentage failed (47%) followed by October (43%). To enhance this chart even further we could filter on goals with about $10,000 since Louise is budgeting over $10,000 for her campaign.
