# VBA Homework - The VBA of Wall Street

## **Project Purpose**

* For this project, a VBA script was created that loops through all the stocks for one year and provides the following outputs per stock:

    * The ticker symbol.

    * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

    * The percent change from opening price at the beginning of a given year to the closing price at the end of that year. If the yearly percent change is positive, the cells are conditionally formatted to appear green. If the yearly percent change is negative, the cells are conditionally formatted to appear red. 

    * The total stock volume of the stock.
    
## **Submited Materials**
* The following was uploaded for this project:
    * Screen-shots showing the script output per year 
    * VBA scripts (Please see **Additional Notes** below)
    
## **Additional Notes**
* Two VBA scripts are included with the submitted materials. The **VBA_Stocks_Single.bas** script works on one single sheet of the workbook at a time. Unfortunately, this was the only way I could get the final outputs on the multi-year workbook of data to run. The other script, **VBA_Stocks_AllSheets.bas**, is supposed to loop through all the sheets of the workbook. This script appeared to work as intended when used in the smaller test workbook that was provided with the assignment materials. In the smaller workbook, the script looped through all the sheets of the workbook in one push. However, I was not able to get this script to work on the multi-year workbook and I'm not really sure why. It gave an error about not dividing by 0 on the line for percent_change = (yearly_change / open price) and unfortunately, I wasn't able to figure out how to resolve the error and make the script work on the final multi-year workbook. Any insights that come up when grading this as to where I went wrong or how I would fix this in the future would be much appreciated.
