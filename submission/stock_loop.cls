VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_loop():

    For iWorksheet = 1 To 4:

        Set ws = Worksheets("Q" & iWorksheet)

        ' FIRST SECTION
        ' --------
        ' column and row titles
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' data
        Dim ticker As String: ticker = ws.Cells(2, 1).Value
        Dim ticker_row As Integer: ticker_row = 2
        Dim ticker_quarterly_open As Double: ticker_quarterly_open = ws.Cells(2, 3).Value
        Dim ticker_quarterly_close As Double: ticker_quarterly_close = -1
        Dim ticker_total_stock_volume As Double: ticker_total_stock_volume = 0
        
        Dim last_row As Long: last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For iRow = 2 To last_row:
        
            ticker_cell = ws.Cells(iRow, 1).Value
        
            If Not ticker = ticker_cell Then
            
                ' calculate
                ticker_quarterly_close = ws.Cells(iRow - 1, 6).Value
                Dim quarterly_change As Double: quarterly_change = ticker_quarterly_close - ticker_quarterly_open
                Dim percent_change As Double: percent_change = (ticker_quarterly_close - ticker_quarterly_open) / ticker_quarterly_open
                ' print results
                ws.Cells(ticker_row, 9).Value = ticker
                ws.Cells(ticker_row, 10).Value = quarterly_change
                ws.Cells(ticker_row, 11).Value = percent_change
                ws.Cells(ticker_row, 12).Value = ticker_total_stock_volume
                
                ' update variables
                ticker = ticker_cell
                ticker_row = ticker_row + 1
                ticker_quarterly_open = ws.Cells(iRow, 3).Value
                ticker_total_stock_volume = 0
                
            End If
            
            ticker_total_stock_volume = ticker_total_stock_volume + ws.Cells(iRow, 7).Value
            
        Next iRow
        
        ' conditional formatting
        ws.Columns("J:L").AutoFit
        
        Set quarterly_change_range = ws.Range("J2:J1500")
        quarterly_change_range.FormatConditions.Delete
        quarterly_change_range.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0").Interior.Color = RGB(198, 224, 180)
        quarterly_change_range.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0").Interior.Color = RGB(255, 198, 206)
        
        Set percent_change_range = ws.Range("K2:K1500")
        percent_change_range.FormatConditions.Delete
        percent_change_range.NumberFormat = "0.00%"
        
        ' SECOND SECTION
        ' --------
        'column and row titles
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' calculate
        Set percent_change_range = ws.Range("K2:K1500")

        Dim percent_change_max_value As Double: percent_change_max_value = Application.WorksheetFunction.Max(percent_change_range)
        Dim percent_change_max_row As String: percent_change_max_row = Application.WorksheetFunction.Match(percent_change_max_value, percent_change_range, 0) + 1
        Dim percent_change_max_ticker As String: percent_change_max_ticker = ws.Cells(percent_change_max_row, 9)
        
        Dim percent_change_min_value As Double: percent_change_min_value = Application.WorksheetFunction.Min(percent_change_range)
        Dim percent_change_min_row As String: percent_change_min_row = Application.WorksheetFunction.Match(percent_change_min_value, percent_change_range, 0) + 1
        Dim percent_change_min_ticker As String: percent_change_min_ticker = ws.Cells(percent_change_min_row, 9)
        
        Set total_stock_volume_range = ws.Range("L2:L1500")

        Dim total_stock_volume_max_value As Double: total_stock_volume_max_value = Application.WorksheetFunction.Max(total_stock_volume_range)
        Dim total_stock_volume_max_row As String: total_stock_volume_max_row = Application.WorksheetFunction.Match(total_stock_volume_max_value, total_stock_volume_range, 0) + 1
        Dim total_stock_volume_max_ticker As String: total_stock_volume_max_ticker = ws.Cells(total_stock_volume_max_row, 9)
        
        ' print results
        ws.Cells(2, 16).Value = percent_change_max_ticker
        ws.Cells(2, 17).Value = percent_change_max_value
        
        ws.Cells(3, 16).Value = percent_change_min_ticker
        ws.Cells(3, 17).Value = percent_change_min_value
        
        ws.Cells(4, 16).Value = total_stock_volume_max_ticker
        ws.Cells(4, 17).Value = total_stock_volume_max_value
        
        ' conditional formatting
        ws.Columns("O:Q").AutoFit
        
        Set value_percent_change_range = ws.Range("Q2:Q3")
        value_percent_change_range.FormatConditions.Delete
        value_percent_change_range.NumberFormat = "0.00%"

    Next iWorksheet
    
End Sub


