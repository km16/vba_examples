Attribute VB_Name = "arithmetic"
Option Explicit


Sub calcAdd()

Cells(12, 1).Value = Cells(10, 1) + Cells(11, 1)

End Sub


Sub calcSubtract()

Cells(12, 8).Value = Cells(10, 8) - Cells(11, 8)

End Sub


Sub calcMultiply()

Cells(16, 1).Value = Cells(14, 1) * Cells(15, 1)

End Sub


Sub calcDivide()

Cells(16, 8).Value = Cells(14, 8) / Cells(15, 8)

End Sub


Sub calcWholeDivide()

Cells(20, 1).Value = Cells(18, 1) \ Cells(19, 1)

End Sub


Sub calcModulo()

Cells(20, 8).Value = Cells(18, 8) Mod Cells(19, 8)

End Sub
