VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stocks()

Dim row As Double

Dim ticker_current As String
Dim ticker_first_row As Double
Dim ticker_last_row As Double

Dim percent_change As Double
Dim i As Double
Dim total_vol As Double

Dim great_perc_incr As Double
Dim great_perc_decr As Double
Dim great_total_vol As Double
Dim ticker_incr As String
Dim ticker_decr As String
Dim ticker_vol As String

For Each ws In ThisWorkbook.Worksheets
ws.Activate

'change column width to fit text
ws.Range("A1:Q1").EntireColumn.AutoFit

'find last_row value
last_row = ws.Cells(Rows.Count, 1).End(xlUp).row - 1

'add column names
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

'clear great vals
great_perc_incr = 0
great_perc_decr = 0
great_total_vol = 0

'row of first ticker
ticker_first_row = 2
'set first ticker value
ticker_current = Cells(ticker_first_row, 1).Value
'open val of first ticker
ticker_first_row_open = Cells(ticker_first_row, 3).Value

'count of number of tickers completed
'start at 2 because of headers
ticker_number = 2

For row = 2 To last_row
    
    'find where ticker val changes
    If Cells(row + 1, 1).Value <> ticker_current Then
        
        'populate ticker val
        Cells(ticker_number, 9).Value = ticker_current
        
        'save close value of ticker
        ticker_last_row = row
        ticker_last_row_close = Cells(row, 6).Value
        
        'find yearly change
        'close - open
        yearly_change = ticker_last_row_close - ticker_first_row_open
        Cells(ticker_number, 10).Value = yearly_change
        
        'color code Yearly Change
        If yearly_change > 0 Then
        
            Cells(ticker_number, 10).Interior.ColorIndex = 4
        
        Else
        
            Cells(ticker_number, 10).Interior.ColorIndex = 3
            
        End If
        
        'percent change
        percent_change = yearly_change / ticker_first_row_open
        Cells(ticker_number, 11).Value = percent_change
        
        'update greatest % vals
        If percent_change > great_perc_incr Then
                
            great_perc_incr = percent_change
            ticker_incr = ticker_current
                
        End If
                    
        If percent_change < great_perc_decr Then
                
            great_perc_decr = percent_change
            ticker_decr = ticker_current
                
        End If
        
        'find total stock volume
        'add all vol values of ticker
        For i = ticker_first_row To ticker_last_row
            
            total_vol = total_vol + Cells(i, 7).Value
        
        Next i

        'populate total stock volume
        Cells(ticker_number, 12).Value = total_vol
        
        'update greatest total vol
        If total_vol > great_total_vol Then
                
            great_total_vol = total_vol
            ticker_vol = ticker_current
                
        End If
        
        total_vol = 0
        
        'update to next ticker
        ticker_number = ticker_number + 1
        ticker_first_row = ticker_last_row + 1
        ticker_current = Cells(ticker_first_row, 1).Value
        ticker_first_row_open = Cells(ticker_first_row, 3).Value
        
    End If
        
Next row

'format as percentage
Range("K2:K" & last_row).NumberFormat = "0.00%"
Range("Q2:Q3").NumberFormat = "0.00%"

'set greatest vals
Cells(2, 16).Value = ticker_incr
Cells(3, 16).Value = ticker_decr
Cells(4, 16).Value = ticker_vol
Cells(2, 17).Value = great_perc_incr
Cells(3, 17).Value = great_perc_decr
Cells(4, 17).Value = great_total_vol

Next ws


End Sub

