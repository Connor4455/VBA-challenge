VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub alphabetical_testing()

Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
        lastrow = Cells(Rows.Count, 1).End(xlUp).row
                
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
            
            
        Dim ticker As String
        Dim stock_volume As Double
        stock_volume = 0
        Dim open_price As Double
        Dim close_price As Double
        Dim percent_change As Double
        Dim yearly_change As Double
        Dim i As Long
        Dim row As Double
        row = 2
        Dim col As Integer
        col = 1
        


        
        For i = 2 To lastrow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker = Cells(i, col).Value
                Cells(row, col + 8).Value = ticker
                open_price = Cells(i, col + 2).Value
                close_price = Cells(i, col + 5).Value
                yearly_change = close_price - open_price
                Cells(row, col + 9).Value = yearly_change
                If (open_price = 0 And close_price = 0) Then
                    percent_change = 0
                ElseIf (open_price = 0 And close_price <> 0) Then
                    percent_change = 1
                Else
                    percent_change = yearly_change / open_price
                    Cells(row, col + 10).Value = percent_change
                    Cells(row, col + 10).NumberFormat = "0.00%"
                End If
                stock_volume = stock_volume + Cells(i, col + 6).Value
                Cells(row, col + 11).Value = stock_volume
                row = row + 1
                open_price = Cells(i + 1, col + 2)
                stock_volume = 0
            Else
                stock_volume = stock_volume + Cells(i, col + 6).Value
                
            End If
                
        Next i
        
        yearly_change_last_row = Cells(Rows.Count, 9).End(xlUp).row
      
        For j = 2 To yearly_change_last_row
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 10
            ElseIf Cells(j, col + 9).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
    
    
    Next ws
        
End Sub
