# VBA_Challenge

For this challenge , I used the data provided in the Multiple_year_stock_data excel spreadsheet to complete a comprehensive multi-year analysis from 2018-2020.

## External_Sources 

* All code in the VBA_Challenge Excel Workbook I have attached to this repository is desigined to run on all three of the sheets provided in the workbook.  In order to accomplish this, I used code I found off of https://excelchamps.com/vba/loop-sheets/.

* In order to find the last row in all of the of the excel sheets, I used code I found from https://www.wallstreetmojo.com/vba-last-row/#h-excel-vba-last-row.

## Stock_Analysis 

* In order to calculate the difference between the opening and closing price for a particular stock ticker on a given year, I created a main for-loop to detect when the stock ticker changed in Column A and a inner_loop that would calculate the opening year price in Column C from the closing price in Column G when a new stock ticker is looped through.

* The data collected from the two loops are placed in columns I and J respectively. In addition , two more columns were created; "Perent Yearly Change" and "Total Stock Volume." These calculations were also derived from the two loops and placed in columns K and L respectively.

* Whenever the loop starts over again,  the counter for "Total Stock Volume" always goes back to 0 for the new stock ticker. 

## Colors

* The difference in the open to closing stock prices calculated in Column J is color-coded based on their values.  If the yearly change is > 0 then the cell is color-coded green.  If the yearly change is < 0, the cell is color-coded red.

* The "formatter" activity included in the Day 3 activities for VBA helped me greatly with this part of the challenge.

## Stats

* In Columns O-Q of each worksheet you will find a statistical analysis for all three worksheets.  The greatest and lowest percent change, as well as the highest volume, have been calculated and placed in those columns adjacent to their respective stock ticker.

* In order to identify their respective stock ticker, an inner loop was created in that whenever the statistical datapoints were discovered in their respective columns, the corresponding stock ticker would be identified in its respective column
  
