Attribute VB_Name = "Module1"
Sub Stock_Challenge()

'Notes on code:
'1 - please note all three solutions are contained in this code.
'2 - this code finds the required ticker, change over year, % change and total volume
'3 - finds greatest % increase, greatest % decrease, greatest total volume
'4 - loops through all sheets in workbook.


'variable declarations
'use of double precision required to avoid propogation of summation errors in trade volume
Dim c As Double 'counter for number of tickers
Dim r As Double 'counter for rows number
Dim num_tickers As Double 'number of unique tickers or stocks
Dim open_price() As Double 'opening value beginning of year
Dim close_price() As Double 'closing value end of year
Dim trade_volume() As Double ' total trading volume in one year
Dim ticker_info() As String 'array that holds necessary stock ticker.  Defined as dynamic
Dim currentsheet As String ' placeholder to check sheet name to help trouble shoot

'dimensioning items for bonus
Dim ticker_max_increase As String
Dim max_increase As Double
Dim ticker_max_decrease As String
Dim max_decrease As Double
Dim ticker_max_volume As String
Dim max_volume As Double

Dim ws As Worksheet ' for cycling through multiple sheets (part of bonus)

For Each ws In Worksheets ' cycling through worksheets

    With ws 'applying with executes all commands on that specific sheet
      
      'Find how many rows exist to set looping limits
      Dim last_row As Double
      last_row = .Cells(Rows.Count, 1).End(xlUp).Row
      
      'what is current sheet - to help error checking.  Msgbox's to help display names for trouble shooting.
      'MsgBox ("work sheeet is: " + currentsheet)
      'MsgBox ("number of rows: " + Str(last_row))
        
      num_tickers = 1 'start at one because start counting on first ticker
      
      'Find number of unique tickers - assumes sorted so tickers are in order and not mixed
      For r = 2 To last_row - 1 'loop starting after header then down to 2nd last row
          If .Cells(r, 1).Value <> .Cells(r + 1, 1).Value Then 'comparing current and next row.
              num_tickers = num_tickers + 1 'if rows are different then must be new ticker so increment
          End If
      Next r
      
      'MsgBox ("number of unique tickers: " + Str(num_tickers))
    
        'redimension arrays knowing the number of unique tickers - clears each array for each sheet
        ReDim ticker_info(num_tickers) As String 'Setting number of rows based on num_tickers
        ReDim open_price(num_tickers) As Double 'opening value beginning of year
        ReDim close_price(num_tickers) As Double 'closing value end of year
        ReDim trade_volume(num_tickers) As Double 'for total trade volume (sum)
        
        c = 1 'counter to increment tickers
        r = 2 'row increments, starting at two avoiding the header
        
        For c = 1 To num_tickers 'will increment through all tickers
            
            open_price(c) = .Cells(r, 3).Value ' assigning initial values based on first entry
            close_price(c) = .Cells(r, 6).Value 'assigning initial values based on first entry
            
            '*************************************************
            'loop below cycles through all information for given ticker
            '*************************************************
            Do While .Cells(r, 1).Value = .Cells(r + 1, 1).Value 'comparing to make sure same ticker
                
                ' if current date is greater than next date than assign new opening price - goal to find opening year value
                ' however also need to check for earlier prices that are zero.  Therefore check that next value is not zero
                ' and check if next opening price is zero and old opening is zero.
                If (.Cells(r, 2).Value > .Cells(r + 1, 2).Value And .Cells(r + 1, 3) <> 0) _
                Or (open_price(c) = 0 And .Cells(r + 1, 3) <> 0) Then
                
                    open_price(c) = .Cells(r + 1, 3).Value 'only assign if (i) non-zero, (ii) lower date (both have non zero opening price
                                                            'or (iii) if current open is zero and new open is not zero.
                End If
                
                ' if current date is less than next date than assign new closing - goal to find close of year value
                If .Cells(r, 2).Value < .Cells(r + 1, 2).Value Then
                    close_price(c) = .Cells(r + 1, 6).Value
                End If
                
                trade_volume(c) = trade_volume(c) + .Cells(r, 7).Value 'summing total volume
                r = r + 1 'incrementing row counter
                
            Loop
            
            'summing total volume capture last one (when tickers don't equal on r and r+1 skip summing)
            trade_volume(c) = trade_volume(c) + .Cells(r, 7).Value
        
            ticker_info(c) = .Cells(r, 1).Value  'putting in ticker price
          
            r = r + 1 ' needs to increment into next section because when cell(r,1)<>cell(r+1,1) then does not increment in loop
            
            If r > last_row Then 'to error check and catch when reach end of list
                Exit For
            End If
                
        Next c
        
        'outputting results
        'Headers
        .Cells(1, 9).Value = "Ticker"
        .Cells(1, 10).Value = "Change in Price"
        .Cells(1, 11).Value = "Percent Change"
        .Cells(1, 12).Value = "Trade Volume"
        
        For c = 1 To num_tickers
            .Cells(c + 1, 9).Value = ticker_info(c)
            .Cells(c + 1, 10).Value = close_price(c) - open_price(c)
            
            'Conditional Formatting for change in value over year
            If (close_price(c) - open_price(c)) < 0 Then ' conditional formating for cell color
                .Cells(c + 1, 10).Interior.ColorIndex = 3 ' red for negative
            ElseIf (close_price(c) - open_price(c)) > 0 Then
                .Cells(c + 1, 10).Interior.ColorIndex = 4 'green for good
            Else
                .Cells(c + 1, 10).Interior.ColorIndex = 6 'yellow for no change
            End If
                
            If open_price(c) <> 0 Then 'some stocks have zero opening value - new to market.  Need to capture
                .Cells(c + 1, 11).Value = (close_price(c) / open_price(c)) - 1
                .Cells(c + 1, 11).NumberFormat = "0.00%" 'Setting percen format two decimals
            Else
                .Cells(c + 1, 11).Value = "NA"
            End If
            .Cells(c + 1, 12).Value = trade_volume(c)
        Next c

'***********************************************************
        'Bonus work - finding specific metrics
        'setting all initial values
        ticker_max_increase = ticker_info(1)
        max_increase = close_price(1) / open_price(1) - 1
        ticker_max_decrease = ticker_info(1)
        max_decrease = close_price(1) / open_price(1) - 1
        ticker_max_volume = ticker_info(1)
        max_volume = trade_volume(1)
        
        
        For c = 2 To num_tickers 'running through all lines
            
            If open_price(c) <> 0 Then 'prevent bad calculation
                
                'checking for max increase
                If (close_price(c) / open_price(c) - 1) > max_increase Then
                    ticker_max_increase = ticker_info(c)
                    max_increase = (close_price(c) / open_price(c) - 1)
                End If
                
                'checking for min increase
                If (close_price(c) / open_price(c) - 1) < max_decrease Then
                    ticker_max_decrease = ticker_info(c)
                    max_decrease = (close_price(c) / open_price(c) - 1)
                End If
            End If
            
            'checking for max volume
            If trade_volume(c) > max_volume Then
                ticker_max_volume = ticker_info(c)
                max_volume = trade_volume(c)
            End If
                
        Next c

        'output of bonus work
        'Headers of columns
        .Cells(1, 16).Value = "Ticker"
        .Cells(1, 17).Value = "Value"
        .Cells(2, 15).Value = "Max % Increase"
        .Cells(3, 15).Value = "Max % decrease"
        .Cells(4, 15).Value = "Max Volume "
        
        .Cells(2, 16).Value = ticker_max_increase
        .Cells(2, 17).Value = max_increase
        .Cells(2, 17).NumberFormat = "0.00%" 'Setting percent format two decimals
        
        .Cells(3, 16).Value = ticker_max_decrease
        .Cells(3, 17).Value = max_decrease
        .Cells(3, 17).NumberFormat = "0.00%" 'Setting percen format two decimals
        
        .Cells(4, 16).Value = ticker_max_volume
        .Cells(4, 17).Value = max_volume

        .Range("A:Q").EntireColumn.AutoFit 'auto set column widths so looks nicer
    End With 'Close with for commands with selected sheet
    
    
Next ws ' cycle to next work sheet



End Sub

