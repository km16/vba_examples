Attribute VB_Name = "dataTypes"
Option Explicit


Sub overflowInt()

Dim i As Integer
Dim j As Integer

i = (2 ^ 16) / 2 - 3

For j = 0 To 2

Cells(11, 7).Value = i

Application.Wait (Now + #12:00:01 AM#)

i = i + 1

Next j

End Sub

Sub String1()

Dim x As String

x = "<Text>"

MsgBox x

End Sub

