Attribute VB_Name = "Modul1"
Option Explicit

Sub Nicht()

If Not (1 > 10) Then

MsgBox ("Yeah")

Else

End If
 
End Sub

Sub Und()

If (3 > 1) And (2 < 10) Then

MsgBox ("Yeah")

Else

End If

End Sub

Sub Oder()

If Cells(1, 1).Value > 1 Or Cells(1, 2).Value > 0.5 Then

MsgBox ("Yeah")

Else

End If

End Sub

Sub Oder2()

If Cells(1, 1).Value = 1 Xor Cells(1, 2).Value = 0.5 Then

MsgBox ("Yeah")

Else

End If

End Sub



