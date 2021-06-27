Attribute VB_Name = "Module1"
Sub Stock_VBA()
   
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Volume As Double
    Volume = 0
Dim Ticker_Symbol As String
Dim opening_price As Double
Dim closing_price As Double
Dim Row As Double
    Row = 2
Dim Column As Integer
    Column = 1
Dim i As Long

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
        opening_price = Cells(2, 3).Value
        
        For i = 2 To lastrow
     
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                Yearly_Change = closing_price - opening_price
                Cells(Row, 10).Value = Yearly_Change
                Ticker_Symbol = Cells(i, 1).Value
                Cells(Row, 9).Value = Ticker_Symbol
                closing_price = Cells(i, 6).Value
        
            If (opening_price = 0 And closing_price = 0) Then
                    Percent_Change = 0
                ElseIf (opening_price = 0 And closing_price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / opening_price
                    Cells(Row, 11).Value = Percent_Change
                    Cells(Row, 11).NumberFormat = "0.00%"
                End If
                
        Volume = Volume + Cells(i, 7).Value
        Cells(Row, 12).Value = Volume
                Row = Row + 1
                opening_price = Cells(i + 1, 3)
                Volume = 0
            Else
                Volume = Volume + Cells(i, 7).Value
            End If
        Next i
        
        xyzlastrow = WS.Cells(Rows.Count, 9).End(xlUp).Row
        
        For j = 2 To xyzlastrow
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
    Next WS
        
End Sub


