Attribute VB_Name = "Module1"
Sub sample_test()

Dim WS_Count As Worksheet
 
 For Each WS_Count In ActiveWorkbook.Worksheets
 WS_Count.Activate
  Dim Ticker_value As String
  Dim Volume As Double
  Dim Row As Double
  Dim i As Long
  Dim j As Long
  Dim open_price As Double
  Dim close_price As Double
  Dim Yearly_price As Double
  Dim Percent_Change As Double
  Dim start As Double


   Volume = 0
   Row = 2
   Range("I1").Value = "Ticker Value"
   Range("J1").Value = "Yearly Change"
   Range("K1").Value = "Percent Change"
   Range("L1").Value = "Total Volume"
  
   start = 2
 
   LastRow = Cells(Rows.Count, 1).End(xlUp).Row
 
   For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker_value = Cells(i, 1).Value
        Cells(Row, 9).Value = Ticker_value
        open_price = Cells(start, 3).Value
        close_price = Cells(i, 6).Value
        
        Yearly_price = close_price - open_price
        Cells(Row, 10).Value = Yearly_price
        If open_price = 0 Then
            Percent_Change = 0
        Else
            Percent_Change = Yearly_price / open_price
        End If
        Cells(Row, 11).Value = Percent_Change
        Cells(Row, 11).NumberFormat = "0.00%"
        Volume = Volume + Cells(i, 7).Value
        Cells(Row, 12).Value = Volume
        Row = Row + 1
        start = i + 1
        Volume = 0
    Else
        Volume = Volume + Cells(i, 7).Value
    End If
    
  Next i
 
   LastRowTwo = Cells(Rows.Count, 9).End(xlUp).Row
 
   For i = 2 To LastRowTwo
    If Cells(i, 10).Value > 0 Or Cells(i, 10).Value = 0 Then
        Cells(i, 10).Interior.Color = RGB(0, 255, 0)
    ElseIf Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.Color = RGB(255, 0, 0)
    End If
   Next i
    
   Range("O2").Value = "Greatest Percent Increase"
   Range("O3").Value = "Greatest Percent Decrease"
   Range("O4").Value = "Greatest Total Volume"
   Range("P1").Value = "Ticker"
   Range("Q1").Value = "Value"
   Dim greatest_increase As Double
   Dim greatest_decrease As Double
   Dim greatest_total As Double

 
   For j = 2 To LastRowTwo
    If Cells(j, 11).Value = Application.WorksheetFunction.Max(WS_Count.Range("K2:K" & LastRowTwo)) Then
        Range("P2").Value = Cells(j, 9).Value
        greatest_increase = Cells(j, 11).Value
        Range("Q2").Value = greatest_increase
        Range("Q2").NumberFormat = "0.00%"
    ElseIf Cells(j, 11).Value = Application.WorksheetFunction.Min(WS_Count.Range("K2:K" & LastRowTwo)) Then
        Range("P3").Value = Cells(j, 9).Value
        greatest_decrease = Cells(j, 11).Value
        Range("Q3").Value = greatest_decrease
        Range("Q3").NumberFormat = "0.00%"
    ElseIf Cells(j, 12).Value = Application.WorksheetFunction.Max(WS_Count.Range("L2:L" & LastRowTwo)) Then
        Range("P4").Value = Cells(j, 9).Value
        greatest_total = Cells(j, 12).Value
        Range("Q4").Value = greatest_total
    End If
   Next j
 Next WS_Count
End Sub
