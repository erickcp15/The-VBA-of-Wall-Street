Attribute VB_Name = "Module1"
Sub test3():

Set ws = ActiveSheet

Dim Name As String
Dim Volume As Double
Volume = 0

lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
openp = Cells(2, 3).Value


For i = 2 To lastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        Name = Cells(i, 1).Value
        Volume = Volume + Cells(i, 7).Value
        Yearly_c = Cells(i, 6).Value - openp
        If openp <> 0 Then
            percen = (Yearly_c / openp) * 100
        Else
            percen = 0
        End If


        Range("I" & Summary_Table_Row).Value = Name
        Range("L" & Summary_Table_Row).Value = Volume
        Range("J" & Summary_Table_Row).Value = Yearly_c
        Range("K" & Summary_Table_Row).Value = percen

        Summary_Table_Row = Summary_Table_Row + 1
        Volume = 0
        openp = Cells(i + 1, 3).Value

    Else

        Volume = Volume + Cells(i, 7).Value

    End If

Next i

For i = 2 To lastRow

    If Cells(i, 10) > 0 Then
        Cells(i, 10).Interior.Color = vbGreen
    Else
        Cells(i, 10).Interior.Color = vbRed
    End If

Next i

Dim Max As Double
Dim Min As Double
Dim Grt_V As Double

Max = WorksheetFunction.Max(Range("K:K"))
Max_t = WorksheetFunction.Match(Max, Range("K:K"), 0)
Min = WorksheetFunction.Min(Range("K:K"))
Min_t = WorksheetFunction.Match(Min, Range("K:K"), 0)
Grt_V = WorksheetFunction.Max(Range("L:L"))
V_T = WorksheetFunction.Match(Grt_V, Range("L:L"), 0)

Cells(2, 16).Value = Cells(Max_t, 9)
Cells(3, 16).Value = Cells(Min_t, 9)
Cells(4, 16).Value = Cells(V_T, 9)
Cells(2, 17).Value = Max
Cells(3, 17).Value = Min
Cells(4, 17).Value = Grt_V

'MsgBox (V_T)
'MsgBox (Min_t)
'MsgBox (Max_t)

Columns("I:Q").EntireColumn.AutoFit
Cells(1, 1).Select

End Sub
'Didnt Finish Just the work I did up the point
Sub Didntfinish()

Sheets.Add.Name = "Combined"
Sheets("Combined").Move Before:=Sheets(1)
Set combined_sheet = Worksheets("Combined")

For Each ws In Worksheets

lastRow = combined_sheet.Cells(Row.Count, "A").End(xlUp).Row + 1
lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
combined_sheet.Range("A" & lastRow & ":Q" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:Q" & (lastRowState + 1)).Value

Next ws

combined_sheet.Range("A1:Q1").Value = Sheets(2).Range("A1:Q1").Value

combined_sheet.Columns("A:Q").AutoFit

End Sub

