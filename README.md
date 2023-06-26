# Module-Challenge-2
Challenge Files

Stock Analysis Macro
Summary

The "stock_analysis" VBA macro is designed to analyze stock data in multiple worksheets within a workbook. It calculates various metrics such as yearly change, percent change, and total stock volume for each stock ticker. Additionally, it identifies stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

#Code Explanation

The code uses a loop to iterate through each worksheet in the workbook. Within each worksheet, it processes the stock data row by row to calculate the required metrics. The code initializes variables and sets up the column headers for the results. It determines the row number of the last row with data in column A. Using nested loops, the code iterates through each row of stock data. When a new stock ticker is encountered (i.e., the ticker symbol changes), it calculates and stores the results for the previous stock. If the total stock volume is zero, it handles this special case by displaying zero change, zero percent change, and zero total stock volume. For non-zero total stock volume, it calculates the change and percent change based on the opening and closing prices.

The results are displayed in the appropriate columns, with positive changes colored green, negative changes colored red, and zero changes without color.
After processing all stock data, the code identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. It stores the ticker symbols and corresponding values in separate cells.
The loop moves on to the next worksheet and repeats the process until all worksheets have been analyzed.

#Data Screenshots

![image](https://github.com/M-2499/VBA-challenge/assets/135250810/1fdc243a-2927-41c6-83d1-cc9bf6573028)

![image](https://github.com/M-2499/VBA-challenge/assets/135250810/1ec2a3e8-cfe2-4279-bc9d-9d28ad7aafc2)

![image](https://github.com/M-2499/VBA-challenge/assets/135250810/efa45ab4-f95e-4d43-b40a-144f36a82acb)


#Conclusions

The "stock_analysis" macro provides a useful tool for analyzing stock data across multiple worksheets. It calculates important metrics such as yearly change and percent change, and identifies stocks with the greatest changes and total volumes. The results are displayed in a clear and organized manner, making it easy to analyze and compare different stocks.
With some customization, this macro can be enhanced further to include additional calculations or formatting options based on specific requirements.
Overall, the "stock_analysis" macro can help investors and analysts gain valuable insights from their stock data and make informed decisions.
