

Public Sub Stock()

 Dim ws As Worksheet
 
 Dim i As Double
 Dim Summary_Table_Row As Integer
 Dim Ticker_Total As Double
 
 Dim change As Double
 Dim start As Double
 Dim LastRow As Double
 Dim percentChange As Double

 For Each ws In Worksheets

    Ticker_Total = 0
    Summary_Table_Row = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    start = 2
    change = 0

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
 
 For i = 2 To LastRow
 
    If (ws.Cells(start, 3) = 0 And ws.Cells(i, 6).Value = 0) Then
        percentChange = 0
 
    ElseIf (ws.Cells(start, 3) = 0 And ws.Cells(i, 6) <> 0) Then
       percentChange = 1
 
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

       change = (ws.Cells(i, 6) - ws.Cells(start, 3))
        percentChange = Round((change / ws.Cells(start, 3) * 100), 2)

        ws.Range("I" & Summary_Table_Row).Value = ws.Cells(i, 1).Value
        ws.Range("J" & Summary_Table_Row).Value = Round(change, 2)
        ws.Range("K" & Summary_Table_Row).Value = percentChange & "%"
        ws.Range("L" & Summary_Table_Row).Value = Ticker_Total + ws.Cells(i, 7).Value
    
      ' start of the next stock ticker
        start = i + 1
        
            If (change < 0) Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf (change > 0) Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
    
        Summary_Table_Row = Summary_Table_Row + 1
        Ticker_Total = 0

    Else
    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

    End If

 Next i

    Dim r1 As Range
    Dim r2 As Range

    Set r1 = ws.Range("K2:K" & Rows.Count)
    Set r2 = ws.Range("L2:L" & Rows.Count)
        ws.Range("Q3") = "%" & Application.WorksheetFunction.Min(r1)
        ws.Range("Q2") = "%" & Application.WorksheetFunction.Max(r1)
        ws.Range("Q4") = Application.WorksheetFunction.Max(r2)


        ' find the row of min and max
        a = WorksheetFunction.Match(WorksheetFunction.Max(r1), r1, 0)
        b = WorksheetFunction.Match(WorksheetFunction.Min(r1), r1, 0)
        vol_number = WorksheetFunction.Match(WorksheetFunction.Max(r2), r2, 0)

        ' print the headers
        ws.Range("P2") = ws.Cells(a + 1, 9)
        ws.Range("P3") = ws.Cells(b + 1, 9)
        ws.Range("P4") = ws.Cells(vol_number + 1, 9)
 
 Next ws
 
End Sub



