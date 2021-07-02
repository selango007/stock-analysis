# Analyze Stock using VBA in Excel

# 1. Overview of Project
The overview of this project is to analyze the Stock for all tickers available in the VBA_Analysis.xlsm for 2017 and 2018 years.

# 2. Purpose
The purpose of this project is to help Steve to 
 - Analyze all tickers with the click of the button ver the years and
 - To Improve the performance of the code to execute much quicker.
 
# 3. Results
The results include 2 sections
 - Ticker Performance
 - Code Performance

# 3.1 Ticker Performance
# 3.1.1 Top 2 performing Tickers
1. ticker: ENPH
 - Based on the year 2017 and 2018, this ticker shown good return in both 2017 and 2018.
 - The Total Daily Volume for this ticker increased significantly in 2018 which states that this ticker has a lot of transactions on a day to day basis.
 - Even though the return in 2018 reduced compared to 2017, this ticker is consistently performing well and showing good signs of good performance in the coming years.

2. ticker: RUN
 - This ticker has positive returns in both 2017 and 2018.
 - In 2018, the Total Daily Volume almost doubles when compared to 2017 which shows this ticker is gaining a lot of attraction.
 - This ticker has the out performed all other tickers in terms of returns in 2018.
 - This ticker is definitely a candidate to consider since it is showing consistent improvement over the years.
 
# 3.1.2 Bottom 2 performing Tickers
1. ticker: TERP
 - This ticker had negative returns in both 2017 and 2018. However in 2018 the returns showed a slight improvement.
 - ALso this ticker has shown improvement in 2018 in terms of Total Daily Volume. 
 - It will be interesting to see the performance of this ticker in the coming years.
 
2. ticker: AY
 - This ticker had very less returns in 2017 and in 2018 it dropped to -6.4%
 - There is a significant drop in Total Daily Volume in 2018 when compared to 2017.
 - This shows that the investers are trying to get away from this ticker.

# 3.2 Code Performance
Below is the table comparison of the code performance between Old Code and New Code.

Year    Old Code        New Code
2017    0.82 sec        0.15 sec
2018    0.80 sec        0.15 sec

From the above table, we see there is a 0.7 seconds improvement when the analysis was done using the refactored code.
The code performance is improved by over 80%. This code improvement will certainly help run the analysis much quicker when large amount of data is involved.

# 3.3 Challenges of using this Code
1. This VBA script assumes there are only 12 tickers. If more tickers are to be analysed, then we need to modify the code by increasing the array index.
2. This script will run only for the specific Excel file. Hence the pre-requisite to running this script is to have the data from the desired source loaded to this excel file, so that the analysis can be done.
3. The results generated in the report might not be accurate, since the dtaa is not realtime.
