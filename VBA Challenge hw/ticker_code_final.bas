Attribute VB_Name = "Ticker_code"
Sub Ticker_code()

'Define everything
Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer

'Ignore any wacky divide by zero errors
On Error Resume Next

For Each ws In ThisWorkbook.Worksheets
    'Insert headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
'Setup integers
    Summary_Table_Row = 2
        Row = ActiveSheet.UsedRange.Rows.Count
    
        'Start the loop
        For i = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Get all the values
                 ticker = ws.Cells(i, 1).Value
                   vol = ws.Cells(i, 7).Value
                
                 year_open = ws.Cells(i, 3).Value
                    year_close = ws.Cells(i, 6).Value
        
                 yearly_change = year_close - year_open
                  percent_change = year_close / year_open
           

 'Populate values into summary table
                ws.Cells(Summary_Table_Row, 9).Value = ticker
                ws.Cells(Summary_Table_Row, 10).Value = yearly_change
                ws.Cells(Summary_Table_Row, 11).Value = percent_change
                ws.Cells(Summary_Table_Row, 12).Value = vol
                Summary_Table_Row = Summary_Table_Row + 1
                vol = 0
   
   'Fix open price equal zero issue
        ElseIf open_price <> 0 Then
            price_change_percent = (price_change_percent / open_price) * 100
        
             End If
        Next i
        
    'Format cells to percentage
        ws.Columns("K").NumberFormat = "0.00%"
    
    'Conditional formatting for yearly change
        Dim rg As Range
        Dim g As Long
        Dim c As Long
        Dim color_cell As Range
    
     Set rg = ws.Range("J2", Range("J2").End(xlDown))
        c = rg.Cells.Count
    
    For g = 1 To c
    Set color_cell = rg(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next g
     

'Autofit Columns
    ws.Columns("A:R").EntireColumn.AutoFit

'Find greatest increase, greatest decrease, and greatest total volume
    Range("Q2").Value = WorksheetFunction.Max(Range("K:K"))
                Range("Q2").NumberFormat = "0.00%"
    Range("Q3").Value = WorksheetFunction.Min(Range("K:K"))
                FormatPercent (Range("Q2"))
    Range("Q3").NumberFormat = "0.00%"
                Range("Q4").Value = WorksheetFunction.Max(Range("L:L"))
                                                                   
'Find matching ticker symbols
    For i = 2 To ws.UsedRange.Rows.Count
    
    If Cells(i, 11).Value = Range("Q2").Value Then
            Range("P2").Value = Cells(i, 9).Value
                    End If
                                        
    If Cells(i, 11).Value = Range("Q3").Value Then
            Range("P3").Value = Cells(i, 9).Value
                    End If
                                        
    If Cells(i, 12).Value = Range("Q4").Value Then
            Range("P4").Value = Cells(i, 9).Value
                    End If
                            Next i
                                    Next ws
   
End Sub
