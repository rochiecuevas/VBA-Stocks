## Introduction
The multi-year stock data contains daily information on stock volume, opening and closing values, and the highest and lowest values. The MS Excel workbook contains stock data for three years: 2014, 2015, and 2016. the object of the study was to summarise the stock data per year.

## Method
For this dataset, VBA scripts were written and executed to determine, for each year, the following: total volume per stock, yearly change, and percent change. The stock with the highest total stock volume, the stock with the greatest percent decrease, and the stock with the greatest percent increase were also indicated.

### Total stock volume per ticker
To get the total stock volume (vol) per ticker, subtotals were calculated using the SumIf function in VBA. In this case, Column A contains the tickers and Column G contains the stock volume for each trading day of each ticker. A new list, containing each unique ticker and its volume subtotal, was put in Columns J–M.

        For i = 2 To LastRow
        ticker = Cells(i, 1).Value
        next_ticker = Cells(i + 1, 1).Value
        
        If ticker <> next_ticker Then
            Cells(position, 10).Value = ticker 
            vol = WorksheetFunction.SumIf(Range("A2:A" & LastRow), ticker, Range("G2:G" & LastRow))
            Cells(position, 13).Value = vol
        End If    

### Comparing stock values between the start and the end of the year
The number of tickers increased from 2014 to 2016. The trading dates are listed in Column B. Ideally, each year started on January 1 (<year>0101) and ended on December 30 (<year>1230). However, the trade of many tickers did not start and end on these days; this situation has led to varying ranges per ticker. Hence, it was important to find the first and last occurrences of each ticker in Column A to determine the range of rows per ticker.
          
        ' Determine the row numbers of first and last entries for the year
        RowFirst = Range("A1:A" & LastRow).Find(What:=ticker, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False).Row
        RowLast = Range("A1:A" & LastRow).Find(What:=ticker, LookAt:=xlWhole, SearchDirection:=xlPrevious, MatchCase:=False).Row

Assuming that trading started on the first time the ticker occurred in Column A, then determining the value of each ticker on the year's first (year_open) and last trading day (year_close) was done using this code:

        ' Determine stock value on first day opening and last day closing
        year_open = Cells(RowFirst, 3).Value
        year_close = Cells(RowLast, 6).Value

Yearly change (yearly_change, Column K) per ticker was calculated based on the opening value on the first day and the closing value on the last day. To ease in differentiating negative and positive changes, each cell of yearly_change values was colour-filled as follows: negative changes were red (3) and positive changes were green (4). Cells with yearly_change = 0 were kept uncoloured.

          yearly_change = year_close - year_open
          
          If yearly_change < 0 Then
                    Cells(position, 11).Interior.ColorIndex = 3
            ElseIf yearly_change > 0 Then
                    Cells(position, 11).Interior.ColorIndex = 4
            End If

Obtaining the percent change (percent_change, Column L) between each ticker's opening and closing values was straightforward. But because there were tickers whose opening values were zero at the start of the year and dividing by zero is disallowed, a conditional was added to print NA for tickers whose year_open = 0. 

          If year_open > 0 Then
            percent = yearly_change / year_open
            Cells(position, 12).Value = percent
            Cells(position, 12).NumberFormat = "0.00%"
        ElseIf year_open <= 0 Then
            Cells(position, 12).Value = "NA"
        End If

### Greatest changes
To find the tickers with the greatest changes from Column L for percent change and from Column M for total stock volume, VBA's Application.WorksheetFunction.Max and Application.WorksheetFunction.Min were used. These are equivalent to the Max and Min functions in MS Excel. The following variables were determined:

        gt_in = greatest percent increase
        gt_dc = greatest percent decrease
        gt_vol = greatest total stock volume

The Max and Min functions were determined using the following VBA scripts:

        gt_in = Application.WorksheetFunction.Max(Columns("L"))
        gt_dc = Application.WorksheetFunction.Min(Columns("L"))
        gt_vol = Application.WorksheetFunction.Max(Columns("M"))

To determine the tickers corresponding to these greatest changes, a new loop was conducted; this time, using the unique tickers in Column J. Because this is a new list, it was important to determine the last row with content in this column.

        ' Find the last row
        LastRow = Cells(Rows.Count, 10).End(xlUp).Row

        ' Loop through the column to find the corresponding ticker
        For j = 1 To LastRow

        ' Find the corresponding ticker
        If Cells(j, 12).Value = gt_in Then
           Cells(2, 17).Value = Cells(j, 10).Value
        End If
    
        If Cells(j, 12).Value = gt_dc Then
           Cells(3, 17).Value = Cells(j, 10).Value
        End If
    
        If Cells(j, 13).Value = gt_vol Then
           Cells(4, 17).Value = Cells(j, 10).Value
        End If
    
    Next j

The complete VBA scripts were provided as a .docx file

## Results
In the 2014 dataset, there were 2835 tickers. A majority of these tickers (n = 1714) performed well by having a higher value at the close of the last trading day than the value at the start of the first trading day of the year. Two of the tickers did not show a change in value while 1118 tickers had negative changes between the year's open and the year's close. One ticker, PLNT, did not appear to have been traded in 2014 and it's opening value was set at zero. It received a non-zero valuation and started getting traded on August 6, 2015.

![2014](https://github.com/rochiecuevas/VBA-Stocks/blob/master/VBA_moderate-difficult_2014.png)
Figure 1. Screenshot of the 2014 results


In 2015, 3004 tickers were included in the dataset. Results showed that 1102 of these tickers had higher values at year's end than when they were traded at the start of the year. Eight of the tickers started and ended the year with the same value. On the other hand, 63% of the tickers (n = 1894) had lower values at year's close than when trading started at the year's beginning.

![2015](https://github.com/rochiecuevas/VBA-Stocks/blob/master/VBA_moderate-difficult_2015.png)
Figure 2. Screenshot of the 2015 results


In 2016, 3168 tickers were monitored. At year's end, it was determined that 1935 of these stocks had higher values than when the year started; two tickers did not demonstrate changes; and 1231 tickers reported losses compared to the stock value at the beginning of the year.

![2016](https://github.com/rochiecuevas/VBA-Stocks/blob/master/VBA_moderate-difficult_2016.png)
Figure 3. Screenshot of the 2016 results

For the three consecutive years, BAC consistently had the highest total stock volume among the stocks included in the dataset (Figures 1–3). For 2014 and 2015, the value of BAC was higher at the close of the year than at its opening; for 2015, however, there was a 5.93% decrease in stock value between its opening value for the year and the year's closing value.
