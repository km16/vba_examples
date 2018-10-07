Attribute VB_Name = "logic"
Option Explicit


Sub generateInput()

Cells(9, 2).Value = Application.WorksheetFunction.RoundDown(((1 - 0 + 1) * Rnd + 0), 0)
Cells(11, 2).Value = Application.WorksheetFunction.RoundDown(((1 - 0 + 1) * Rnd + 0), 0)

End Sub


Sub logicNot()

If Not (Cells(9, 2) = 1 Xor Cells(11, 2) = 1) Then

Cells(10, 5).Value = True

Else

Cells(10, 5).Value = False

End If

End Sub


Sub logicAnd()

If (Cells(9, 2) = 1 And Cells(11, 2) = 1) Then

Cells(10, 5).Value = True

Else

Cells(10, 5).Value = False

End If

End Sub


Sub logicOr()

If (Cells(9, 2) = 1 Or Cells(11, 2) = 1) Then

Cells(10, 5).Value = True

Else

Cells(10, 5).Value = False

End If

End Sub


Sub logicXor()

If (Cells(9, 2) = 1 Xor Cells(11, 2) = 1) Then

Cells(10, 5).Value = True

Else

Cells(10, 5).Value = False

End If

End Sub


Sub checkPatient()

Dim station As String
Dim timeSpent As Integer
Dim movePatient As Boolean

station = Cells(17, 10)
timeSpent = Cells(17, 11)

If station = "INTENSIV" And timeSpent > 6 Then

movePatient = True
Cells(17, 12).Value = movePatient

Else

movePatient = False
Cells(17, 12).Value = movePatient

End If

End Sub
