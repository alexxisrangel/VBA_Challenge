Sub stock_analysis()

    ' Define variables
    Dim ws As Worksheet
    Dim last_row As Long
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim summary_table_row_index As Long
    Dim rng As Range
    
     
    
    'Variables for part 2 calculations
    Dim greatest_increase_ticker As String
    Dim greatest_increase As Double
    greatest_increase = 0
    
    Dim greatest_decrease_ticker As String
    Dim greatest_decrease As Double
    greatest_decrease = 0
    
    Dim greatest_volume_ticker As String
    Dim greatest_volume As Double
    greatest_volume = 0
    
    ' Loop through all worksheets in the workbook
    For Each ws In ActiveWorkbook.Worksheets
    
        ' Initialize summary table row index
        summaryTableRowIndex = 2
        
        ' Column headers for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
        ' Find the last row of data for the current worksheet
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all rows of data for the current worksheet
        For i = 2 To last_row
        
            ' Check if the current row is the first row of a new ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                ' Set the ticker symbol
                ticker = ws.Cells(i, 1).Value
                
                ' Set the opening price
                open_price = ws.Cells(i, 3).Value
                
            End If
            
            ' Check if the current row is the last row of a ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                ' Set the closing price
                close_price = ws.Cells(i, 6).Value
                
                ' Calculate the yearly change and percent change
                yearly_change = close_price - open_price
                percent_change = yearly_change / open_price
                
                ' Setting total volume
                total_volume = WorksheetFunction.Sum(ws.Range(ws.Cells(i - 11, 7), ws.Cells(i, 7)))
                
                ' Add a new row to the summary table with the calculated values
                ws.Range("I" & summaryTableRowIndex).Value = ticker
                ws.Range("J" & summaryTableRowIndex).Value = yearly_change
                ws.Range("K" & summaryTableRowIndex).Value = percent_change
                ws.Range("K" & summaryTableRowIndex).NumberFormat = "0.00%"
                ws.Range("L" & summaryTableRowIndex).Value = total_volume
                
                'Conditional format For this I had to use the record macros feature
                 Set rng = ws.Range("K2:K" & summaryTableRowIndex)
                rng.FormatConditions.Delete
                With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
                    .Interior.Color = RGB(146, 208, 80) '
                End With
                With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                    .Interior.Color = RGB(255, 0, 0)
                End With
                
                 Set rng = ws.Range("J2:K" & summaryTableRowIndex)
                rng.FormatConditions.Delete
                With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
                    .Interior.Color = RGB(146, 208, 80)
                End With
                With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                    .Interior.Color = RGB(255, 0, 0)
                End With
                
                'checking for greatest increase/decrease/volume
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                    greatest_increase_ticker = ticker
                End If
                
                If percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                    greatest_decrease_ticker = ticker
                End If
                
                If total_volume > greatest_volume Then
                    greatest_volume = total_volume
                    greatest_volume_ticker = ticker
                End If
                
                'reset variables for new ticker
                ws.Cells(i, 1).Value = ticker
                ws.Cells(i, 3).Value = opening_price
                ws.Cells(i, 6).Value = closing_price
                ws.Cells(i, 7).Value = summary_table_row
                
                ' Increment the summary table row index
                summaryTableRowIndex = summaryTableRowIndex + 1
                
                Else
                    
                    'adding totla volume and update closing price
                    total_volume = total_volume + ws.Cells(i, 7).Value
                    closing_price = ws.Cells(i, 6).Value
                
            End If
       
       'summary table for greatest Increase/decrease/volume
        ws.Cells(2, 14).Value = "Greatest % increase"
        ws.Range("O2").Value = greatest_increase_ticker
        ws.Range("P2").Value = greatest_increase
        ws.Range("O2").NumberFormat = "0.00%"
        
        ws.Cells(3, 14).Value = "Greatest % decrease"
        ws.Range("O3").Value = greatest_decrease_ticker
        ws.Range("P3").Value = greatest_decrease
        ws.Range("P3").NumberFormat = "0.00%"
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Range("O4").Value = greatest_volume_ticker
        ws.Range("P4").Value = greatest_volume
        
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"

   Next i
        
 Next ws
 
 
 End Sub
 


