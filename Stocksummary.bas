Attribute VB_Name = "Module4"
Sub StockSummary()
 
' Note data is already sorted by date in ascending order (earliest to latest)
' Sanitisation to sort data by date and ticker not required

' Define variables
Dim ticker As String
Dim openingprice As Double
Dim closingprice As Double
Dim totalvolume As Double
Dim yearchange As Double
Dim percentchange As Double
Dim ws As Worksheet
Dim lastrow As Long
Dim summaryrow As Long
Dim i As Long

' Loop through all worksheets
For Each ws In Worksheets

' Count the last row of data
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
' Set up column headers to display the summary output
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
  
' Set up initial row for the summary output
    summaryrow = 2
    
' Set up the first ticker and opening price values for comparison
    ticker = ws.Cells(2, 1).Value
    openingprice = ws.Cells(2, 3).Value
    
' Loop through all rows of data
    For i = 2 To lastrow
    
' Check if the ticker in the next row has changed
        If ws.Cells(i, 1).Value <> ticker Then
        
' Set up formula to calculate the year change and percent change
            yearchange = closingprice - openingprice
            
            If openingprice <> 0 Then
                percentchange = yearchange / openingprice

' Set up formula if there is no year change and percent change
            Else
                percentchange = 0
            End If
            
' Display summary information under the column headers
            ws.Range("I" & summaryrow).Value = ticker
            ws.Range("J" & summaryrow).Value = yearchange
            ws.Range("K" & summaryrow).Value = percentchange
            ws.Range("L" & summaryrow).Value = totalvolume

' Set up to display in the next row in the summary columns
            summaryrow = summaryrow + 1
            
' Reset variables for the next ticker
            ticker = ws.Cells(i, 1).Value
            openingprice = ws.Cells(i, 3).Value
                        
' Set up stock volumes to start count number at 0
            totalvolume = 0
            
        End If
        
' Add stock volume for the current row to the ticker
        totalvolume = totalvolume + ws.Cells(i, 7).Value
        
' Set up the closing price for the current ticker
        closingprice = ws.Cells(i, 6).Value
        
' Apply conditional formatting highlight to J column green if positive or red if negative
        If yearchange >= 0 Then
            ws.Range("J" & summaryrow).Interior.ColorIndex = 3
        Else
            ws.Range("J" & summaryrow).Interior.ColorIndex = 4
        End If
    
    Next i
    
' -------------------------------------------------------------------

' Set up formula to calculate the year change and percent change for the last ticker

    yearchange = closingprice - openingprice
    
    If openingprice <> 0 Then
        percentchange = yearchange / openingprice
    Else
        percentchange = 0
    End If
    
' Display summary information for last ticker under the column headers
    ws.Range("I" & summaryrow).Value = ticker
    ws.Range("J" & summaryrow).Value = yearchange
    ws.Range("K" & summaryrow).Value = percentchange
    ws.Range("L" & summaryrow).Value = totalvolume

' Apply conditional formatting highlight to J column green if positive or red if negative for last ticker

    If yearchange >= 0 Then
        ws.Range("J" & summaryrow).Interior.ColorIndex = 3
    Else
        ws.Range("J" & summaryrow).Interior.ColorIndex = 4

    End If
    
' -------------------------------------------------------------------

' Format summary column K to percentage to 2DP

ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"

' Format summary columns to autodit

ws.Columns("I:L").AutoFit


Next ws

End Sub


