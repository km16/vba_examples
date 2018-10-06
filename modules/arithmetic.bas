Attribute VB_Name = "arithmetic"
Option Explicit


Sub calcAdd()

Cells(12, 2).Value = Cells(10, 2) + Cells(11, 2)

End Sub


Sub calcSubtract()

Cells(12, 9).Value = Cells(10, 9) - Cells(11, 9)

End Sub


Sub calcMultiply()

Cells(16, 2).Value = Cells(14, 2) * Cells(15, 2)

End Sub


Sub calcDivide()

Cells(16, 9).Value = Cells(14, 9) / Cells(15, 9)

End Sub


Sub calcWholeDivide()

Cells(20, 2).Value = Cells(18, 2) \ Cells(19, 2)

End Sub


Sub calcModulo()

Cells(20, 9).Value = Cells(18, 9) Mod Cells(19, 9)

End Sub
