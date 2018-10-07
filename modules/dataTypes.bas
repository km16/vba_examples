Attribute VB_Name = "dataTypes"
Option Explicit


Sub overflowInt()

Dim i As Integer
Dim j As Integer

i = 32765

For j = 0 To 2

Cells(11, 7).Value = i

Application.Wait (Now + #12:00:01 AM#)

i = i + 1

Next j

End Sub


Sub addIntToString()

Dim str As String

str = Cells(11, 10)

str = str + "1"

Cells(11, 10).Value = str

End Sub


Sub addCharToString()

Dim str As String

str = Cells(11, 10)

str = str + "a"

Cells(11, 10).Value = str
End Sub
