## Introduction
The multi-year stock data contains daily information on stock volume, opening and closing values, and the highest and lowest values. The MS Excel workbook contains stock data for three years: 2014, 2015, and 2016. the object of the study was to summarise the stock data per year.

## Method
For this dataset, VBA scripts were written and executed to determine, for each year, the following: total volume per stock, yearly change, and percent change. The stock with the highest total stock volume, the stock with the greatest percent decrease, and the stock with the greatest percent increase were also indicated.

The following equations were used to calculate some of the variables:

          yearly change = value of stock at year close - value of stock at year open
          percent change = (yearly change / value of stock at year open) * 100

The VBA scripts were provided as a .docx file

## Results
The number of tickers increased from 2014 to 2016. For the three consecutive years, BAC consistently had the highest total stock volume among the stocks included in the dataset (Figures 1–3). For 2014 and 2015, the value of BAC was higher at the close of the year than at its opening; for 2015, however, there was a 5.93% decrease in stock value between its opening value for the year and the year's closing value.

In the 2014 dataset, there were 2835 tickers. A majority of these tickers (n = 1714) performed well by having a higher value at the close of the last trading day than the value at the start of the first trading day of the year. Two of the tickers did not show a change in value while 1118 tickers had negative changes between the year's open and the year's close. One ticker, PLNT, did not appear to have been traded in 2014. It started getting traded on August 6, 2015.

![2014](https://github.com/rochiecuevas/VBA-Stocks/blob/master/VBA_moderate-difficult_2014.png)
Figure 1. Screenshot of the 2014 results


In 2015, 3004 tickers were included in the dataset. Results showed that 1102 of these tickers had higher values at year's end than when they were traded at the start of the year. Eight of the tickers started and ended the year with the same value. On the other hand, 63% of the tickers(n = 1894) had lower values at year's close than when trading started at the year's beginning.

![2015](https://github.com/rochiecuevas/VBA-Stocks/blob/master/VBA_moderate-difficult_2015.png)
Figure 2. Screenshot of the 2015 results


In 2016, 3168 tickers were monitored. At year's end, it was determined that 1935 of these stocks had higher values than when the year started; two tickers did not demonstrate changes; and 1231 tickers reported losses compared to the stock value at the beginning of the year.

![2016](https://github.com/rochiecuevas/VBA-Stocks/blob/master/VBA_moderate-difficult_2016.png)
Figure 3. Screenshot of the 2016 results
