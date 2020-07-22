Attribute VB_Name = "Module1"
Sub stock_analysis()



    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
    Dim ticker As String
    Dim open_year As Double
    Dim open_value As Double
    open_value = 2
    Dim percent_change As Double
    Dim close_year As Double
    Dim total_ticker As Double
    total_ticker = 0
    Dim OutPut_Row As Integer
    OutPut_Row = 2
    
    
    Cells(1, "I").Value = "Ticker"
    
    Cells(1, "J").Value = "Yearly Change"
    
    Cells(1, "K").Value = "Percent Change"
    
    Cells(1, "L").Value = "Total Stock Volume"
    

    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To lastrow
            open_year = ws.Cells(open_value, 3).Value
            
        
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ticker = ws.Cells(i, 1).Value
        
        total_ticker = total_ticker + ws.Cells(i, 7).Value
       
        close_year = ws.Cells(i, 6).Value
        
        yearly_change = close_year - open_year
        ws.Cells(i, 10).Value = yearly_change

        
        If open_year = 0 Then
            percent_change = 0
        Else
            percent_change = yearly_change / open_year
        End If
       
        ws.Range("K" & OutPut_Row).Value = (percent_change & "%")
        
        ws.Range("J" & OutPut_Row).Value = yearly_change
        
        ws.Range("I" & OutPut_Row).Value = ticker
        
        ws.Range("L" & OutPut_Row).Value = total_ticker
    
    
        If ws.Cells(OutPut_Row, 10).Value > 0 Then
            ws.Cells(OutPut_Row, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(OutPut_Row, 10).Interior.ColorIndex = 3
        End If
        
        
        OutPut_Row = OutPut_Row + 1
        
        total_ticker = 0
        open_value = (i + 1)
        
   
    Else
        total_ticker = total_ticker + ws.Cells(i, 7).Value
        
    End If
    
  Next i

Next ws


End Sub




