VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Multiple_year_stock_data()

Dim ws As Worksheet

For Each ws In Worksheets
    
    'Declare variables
    Dim last_row As Long
    Dim ticker As String
    Dim total_volume As Double
    Dim output_row As Integer
    Dim column As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim last_row_yc As Long

    'Headings
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Initiate variables
    last_row = Cells(Rows.Count, 1).End(xlUp).row
    output_row = 2
    total_volume = 0
    column = 1
    open_price = Cells(2, 3)
    
    'Formatting
    ws.Columns("I:Q").AutoFit
    
    'Loop through tickers
    
    For i = 2 To last_row
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            ticker = ws.Cells(i, 1).Value
        
            total_volume = total_volume + ws.Cells(i, 7).Value
        
            ws.Range("I" & output_row).Value = ticker
        
            ws.Range("L" & output_row).Value = total_volume
            
            close_price = ws.Cells(i, 6).Value
            
            yearly_change = close_price - open_price
            
            ws.Range("J" & output_row).Value = yearly_change
            
            If (open_price = 0 And close_price = 0) Then
            
                percent_change = 0
                
            ElseIf (open_price = 0 And close_price <> 0) Then
                
                percent_change = 1
                
            Else
            
                percent_change = yearly_change / open_price
            
            End If
            
            ws.Range("K" & output_row).Value = percent_change
            
            ws.Range("K" & output_row).NumberFormat = "0.00%"
            
            total_volume = 0
            
            open_price = ws.Cells(i + 1, 3)
            
            output_row = output_row + 1
    
        Else
    
            total_volume = total_volume + ws.Cells(i, 7).Value
        
        End If
    
    Next i

    last_row_yc = ws.Cells(Rows.Count, 10).End(xlUp).row
    
    'Loop for conditional formatting
    
    For j = 2 To last_row_yc
        
     If ws.Cells(j, 10).Value >= 0 Then
     
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
            Else
            
                ws.Cells(j, 10).Interior.ColorIndex = 3
            
        End If
        
    Next j

Next ws

End Sub



