Attribute VB_Name = "Module1"
' This script will loop through all stocks for one year and output:
' The singular Ticker symbol,
' The yearly change from opening price at the beginning of given year to
' closing price at end of year (first opening - last closing),
' The percent change from opening price at beginning of given year to the
' closing price at end of year (percent of first op. - last close.),
' The total (sum) stock volume per stock ticker

Sub stockLooper():

   'Define variables to loop through all worksheets in workbook
   Dim ws As Worksheet

 'loop through all worksheets
 For Each ws In Worksheets:
    ws.Activate

    'setting headers for new columns of worksheet
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'setting initial variable for holding the ticker
    Dim Ticker_Name As String
    
    'setting initial variable for holding total volume per ticker
    Dim Volume_Total As LongLong
    Volume_Total = 0
    
    'setting initial variable for first opening of year
    Dim First_Open As Double
    First_Open = Cells(2, 3).Value
    
    'setting initial variable for last closing of year
    Dim Last_Close As Double
    
    'setting variable for Yearly Change
    Dim Yearly_Change As Double
    
    'setting variable for Percent Change
    Dim Percent_Change As Double
    
    'keep track of location for each ticker in summary columns
    Dim Summary_Column_Row As Integer
    Summary_Column_Row = 2
    
    'Define variable for finding last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop through all stocks
    For I = 2 To LastRow
    
      'check if we are still within same ticker, if it is not...
      If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
      
        'set the ticker name
        Ticker_Name = Cells(I, 1).Value
        
        'add to the volume total
        Volume_Total = Volume_Total + Cells(I, 7).Value
        
        'set the first open for each ticker
        'First_Open = Cells(i + 1, 3).Value
        
        'set the last closing for each ticker
        Last_Close = Cells(I, 6).Value
        
        'set the Yearly Change for each ticker
        Yearly_Change = First_Open - Last_Close
        
        'set the Percent Change for each ticker
        Percent_Change = Yearly_Change / First_Open
        
        'print the Yearly Change in the summary columns
        Range("J" & Summary_Column_Row).Value = Yearly_Change
        
        'print the Percent Change in the summary columns
        'and format as a percent
        Range("K" & Summary_Column_Row).Value = FormatPercent(Percent_Change)
        
        'print the ticker in the summary columns
        Range("I" & Summary_Column_Row).Value = Ticker_Name
        
        'print volume total to summary columns
        Range("L" & Summary_Column_Row).Value = Volume_Total
        
        'add one to the summary column row
        Summary_Column_Row = Summary_Column_Row + 1
        
        'reset the volume total
        Volume_Total = 0
        
        First_Open = Cells(I + 1, 3).Value
        
        'if cell immediately folowing a row is the same ticker...
      Else
        
        'add to the volume total
        Volume_Total = Volume_Total + Cells(I, 7).Value
      
      End If
      
    Next I
    
  Next ws
    
    'outer loop to loop through all worksheets
 For Each ws In Worksheets:
    ws.Activate
    
    'Autofit all cells in Summary Columns
    Range("I:L").EntireColumn.AutoFit
    
    'format column J (Yearly Change) to shade green when positive change, red when negative change, and
    'yellow when no net change (zero)
      'inner loop to loop through all YearlyChanges
      For x = 2 To LastRow
      
        'if cells in J are greater than zero, color green
        If Cells(x, 10) > 0 Then
           Cells(x, 10).Interior.ColorIndex = 4
            
        'if cells in J are less than zero, color red
        ElseIf Cells(x, 10) < 0 Then
            Cells(x, 10).Interior.ColorIndex = 3
               
      End If
      
    Next x

  Next ws
    














End Sub

