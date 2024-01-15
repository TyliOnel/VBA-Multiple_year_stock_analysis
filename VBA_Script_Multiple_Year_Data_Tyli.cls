VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub multiple_year_stock()

'Declare all variables for code
     'Integer variable to show single number value with no decimal
     'Double variable to show larger number value and allows for decimals
     'String variable to show characters themselves in contiguous sequence rather than numeric values, resulting in text
     'As worksheet to allow code to run through multiple worksheets
     'As long to find bottom of sheet due to the size

Dim ws As Worksheet
Dim ticker As String
Dim lastrowA As Long
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
For Each ws In Worksheets

'Set open price to zero
Dim price_row As Long
price_row = 2

'start total stock volume as zero
Total = 0

'Assign summary table and name columns
Dim summary_table_row As Integer
summary_table_row = 2

'Add header to new columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Determine last row in first column - ticker
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'start loop
For i = 2 To lastrow

   'Find all different ticker names
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'Find total stock volume
        ws.Range("L" & summary_table_row).Value = Total

        'Add ticker data
        ws.Cells(summary_table_row, "I").Value = ws.Cells(i, "A").Value

        'Calculate yearly change and percent change
        open_price = ws.Range("C" & price_row).Value
        close_price = ws.Range("F" & i).Value
        yearly_change = close_price - open_price
    
        'if numbers are equaling zero statement
        If open_price = 0 Then
        percent_change = yearly_change
    
        'set else statement
        Else
        percent_change = yearly_change / open_price
        
        'format column I into dollar value
        ws.Range("I").NumberFormat = "$0.00"

        End If
    
     'Next part of the summary data: add values of yearly change and percent change
     ws.Range("J" & summary_table_row).Value = yearly_change
     ws.Range("K" & summary_table_row).Value = percent_change
     
     'format percentage
     ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                  
    'set colours for positive (green) and negative (red)- use conditional formating
    If ws.Range("J" & summary_table_row).Value > 0 Then
    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
         Else
             ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
         End If
   
   'in the summary table add price row
   summary_table_row = summary_table_row + 1
   price_row = i + 1

   'reset the total stock volume in order to add new information
   Total = 0
    
    'if not last row, take total volume for this row and add to overall total volume
    Else
    Total = Total + ws.Range("G" & i).Value
    
    End If

Next i

    'set values to find the greatest percent change
    Greatest_Increase = ws.Range("K2").Value
    Greatest_Decrease = ws.Range("K2").Value
    Greatest_Total = ws.Range("L2").Value
    
    'Determine last row in ticker column
    Lastrow_Ticker = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'Loop through ticker column to find greatest increase percent change
     For j = 2 To Lastrow_Ticker:
           If ws.Range("K" & j + 1).Value > Greatest_Increase Then
              Greatest_Increase = ws.Range("K" & j + 1).Value
              Greatest_Increase_Ticker = ws.Range("I" & j + 1).Value
              
           'find greatest decrease percent change
           ElseIf ws.Range("K" & j + 1).Value < Greatest_Decrease Then
              Greatest_Decrease = ws.Range("K" & j + 1).Value
              Greatest_Decrease_Ticker = ws.Range("I" & j + 1).Value
            
            'find greatest total percent change
            ElseIf ws.Range("L" & j + 1).Value > Greatest_Total Then
              Greatest_Total = ws.Range("L" & j + 1).Value
              Greatest_Total_Ticker = ws.Range("I" & j + 1).Value
              
            End If
            
        Next j
            
            'fill cells with data of greatest increase, decrease and total - names
            ws.Range("P2").Value = Greatest_Increase_Ticker
            ws.Range("P3").Value = Greatest_Decrease_Ticker
            ws.Range("P4").Value = Greatest_Total_Ticker
            
            'fill cells with data of greatest increase, decrease and total - values
            ws.Range("Q2").Value = Greatest_Increase
            ws.Range("Q3").Value = Greatest_Decrease
            ws.Range("Q4").Value = Greatest_Total
        
        'format into percentage
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Auto size columns
    ws.UsedRange.EntireColumn.AutoFit
            
    'fill in all worksheets
    Next ws

End Sub
