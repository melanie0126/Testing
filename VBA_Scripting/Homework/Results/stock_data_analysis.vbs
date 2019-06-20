Sub WorksheetLoop()
Dim ws As Worksheet
    Call reset
' loop through all worksheets
For Each ws In ThisWorkbook.Worksheets

    Call Stock_Data_Analysis(ws)
Next ws

End Sub


Sub Stock_Data_Analysis(ws)
Dim i As Long
Dim j As Long
Dim m As Long
Dim k As Long
Dim ticker() As Variant
Dim init_next As Double
Dim init_previous As Double
Dim Last As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_volume() As Variant
percent_change_max = 0
percent_change_min = 0
total_volume_max = 0
j = 0
k = 1
init_next = 0
init_previous = ws.Cells(2, 3).Value
Last = 0
ReDim ticker(0)
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"
ws.Activate

Rows_Count = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To Rows_Count
    Columns("I:I").Select
    T_value = ws.Cells(i, 1).Value
    'search ticker(column A), if already exists in column I
    Set cell = Selection.Find(What:=T_value, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    r = ActiveCell.Row + 1 'new ticker row position
    c = ActiveCell.Column
    'a new ticker
    If cell Is Nothing Then
    k = k + 1
    ws.Cells(k, 9).Value = ws.Cells(i, 1).Value
    ws.Cells(k, 12).Value = ws.Cells(i, 7).Value 'the new ticker found, put the first vol value into Total volumn
        If (i > 2) Then
            init_next = ws.Cells(i, 3).Value  'assumed the ticker is sorted, the first one of the new ticker would be the opening price of the beginning of the year
            Last = ws.Cells(i - 1, 6).Value ' one row before the beginning of the new ticker should be last day of the previous ticker
            yearly_change = Last - init_previous
            ws.Cells(k - 1, 10).Value = yearly_change
                If (yearly_change >= 0) Then
                    With ws.Cells(k - 1, 10).Interior
                    .Color = RGB(0, 255, 0) 'Green
                    End With
                Else
                    With ws.Cells(k - 1, 10).Interior
                    .Color = RGB(255, 0, 0) 'Red
                    End With
                End If
            If (init_previous <> 0) Then
                percent_change = (yearly_change / init_previous)
                ws.Cells(k - 1, 11).Value = percent_change
                ws.Cells(k - 1, 11).NumberFormat = "0.00%"
            Else
                ws.Cells(k - 1, 11).Value = Null
            End If
            init_previous = init_next
        End If
        If (percent_change > percent_change_max) Then
            percent_change_max = percent_change
            PCM_ticker_max = ws.Cells(k - 1, 9).Value
        End If
        If (percent_change < percent_change_min) Then
            percent_change_min = percent_change
            PCM_ticker_min = ws.Cells(k - 1, 9).Value
        End If
    ReDim Preserve total_volume(k)
    total_volume(k) = ws.Cells(i, 7).Value
    'same ticker
    Else
    total_volume(k) = total_volume(k) + ws.Cells(i, 7).Value
    ws.Cells(k, 12) = total_volume(k)
    'calculating yearly change and percent change for last ticker
        If (i = Rows_Count) Then
            Last = ws.Cells(i, 6).Value
            yearly_change = Last - init_previous
            percent_change = (yearly_change / init_previous)
            ws.Cells(k, 10).Value = yearly_change
            ws.Cells(k, 11).Value = percent_change
        End If
    End If
  Next i
  ws.Range("M1").Value = "Greatest % increase"
  ws.Range("M2").Value = percent_change_max
  ws.Range("M2").NumberFormat = "0.00%"
  ws.Range("N2").Value = PCM_ticker_max
  ws.Range("M3").Value = "Greatest % Decrease"
  ws.Range("M4").Value = percent_change_min
  ws.Range("M4").NumberFormat = "0.00%"
  ws.Range("N4").Value = PCM_ticker_min
  ws.Range("M5").Value = "Greatest total volume"
  ws.Range("M6").Value = Application.WorksheetFunction.Max(total_volume)
  'GTV_row = Application.WorksheetFunction.Max(total_volume)
  'Range("N6").Value = Cells(GTV_row, 1).Value
  
  
End Sub
