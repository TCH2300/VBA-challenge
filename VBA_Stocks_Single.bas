Attribute VB_Name = "Module1"
Sub Stock_homework():
        'directions: create script that will loop through all the stocks for one year and output:
        'Ticker Symbol
        'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
        'Percent change from opening price at the beginning of a given year to the closing price at the end of that year
        'Total stock volume of the stock
        'ALSO: conidtional formatting that will highlight positive change in green and negative change in red
    
    'add headings for summary table
    Cells(1, 9).Value = "Ticker"
        Cells(1, 9).Font.Bold = True
    Cells(1, 10).Value = "Open Price"
        Cells(1, 10).Font.Bold = False
    Cells(1, 11).Value = "Closing Price"
        Cells(1, 11).Font.Bold = False
    Cells(1, 12).Value = "Yearly Change"
        Cells(1, 12).Font.Bold = True
    Cells(1, 13).Value = "Percent Change"
        Cells(1, 13).Font.Bold = True
    Cells(1, 14).Value = "Total Stock Volume"
        Cells(1, 14).Font.Bold = True
    
    'define variables
    Dim ticker_symbol As String
    Dim open_price As Double
    Dim closed_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim volume_total As LongLong
    Dim i As Long
    Dim last_row As Long
    Dim i_summary As Long
    Dim ws As Worksheet
   
    
    'assign variables
    i_summary = 2
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    'iterate through all rows for loop
'    For Each ws In Worksheets
    
        For i = 2 To last_row
            ticker = Cells(i, 1).Value
         
            If Cells(i - 1, 1).Value <> ticker Then
                open_price = Cells(i, 3).Value
                Cells(i_summary, 10).Value = open_price
                    Columns("J").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($*""_""??_);_(@_)"
            End If
            
            If Cells(i, 1).Value = ticker Then
                volume_total = volume_total + Cells(i, 7).Value
            End If
            
            If Cells(i + 1, 1).Value <> ticker Then 'if the ticker has changed, then populate the summary table and increment summary index
                close_price = Cells(i, 6).Value
                Cells(i_summary, 11).Value = close_price
                    Columns("K").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($*""_""??_);_(@_)"
                Cells(i_summary, 9).Value = ticker
                Cells(i_summary, 14).Value = volume_total
                    Columns("N").NumberFormat = "_($* #,##0_);_($* (#,##0);_($*""_""??_);_(@_)"
                volume_total = 0
                yearly_change = close_price - open_price
                Cells(i_summary, 12).Value = yearly_change
                    Columns("L").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($*""_""??_);_(@_)"
                percent_change = (yearly_change / open_price)
                    Columns("M").NumberFormat = "0.00%"
                Cells(i_summary, 13).Value = percent_change
                i_summary = i_summary + 1 'this means everytime the summary table populates, it moves down to the next row.
            End If
            
            If IsNumeric(Cells(i_summary - 1, 13)) = True Then
                If percent_change >= 0 Then
                    Cells(i_summary - 1, 13).Interior.Color = RGB(0, 255, 0)
                Else
                    Cells(i_summary - 1, 13).Interior.Color = RGB(255, 0, 0)
                End If
            End If
    
    Next i

'Next ws

End Sub
