Attribute VB_Name = "Module1"
Sub A():
Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "volume"
Cells(1, 11).Value = "percentage change"
Cells(1, 12).Value = "yearly change"

Dim ticker As String
Dim next_ticker As String
Dim row As Integer
Dim volume As Double
Dim closing As Double
Dim opening As Double
Dim prev_row As Double

prev_row = 1

row = 2
For i = 2 To 753001
ticker = Cells(i, 1).Value
next_ticker = Cells(i + 1, 1).Value
volume = Cells(i, 7).Value + volume

If prev_row <> row Then
opening = Cells(i, 3).Value
prev_row = row

End If


If ticker <> next_ticker Then
Cells(row, 9).Value = ticker
Cells(row, 10).Value = volume
closing = Cells(i, 6).Value
Cells(row, 12).Value = closing - opening
Cells(row, 11).Value = ((closing - opening) / opening) * 100

row = row + 1


volume = 0
End If



Next i


End Sub
