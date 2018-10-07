Attribute VB_Name = "rangeCells"
Option Explicit


Sub selectRangeData()

Range("A11:B25").Select

End Sub


Sub copyRangeData()

If IsEmpty(Cells(11, 7)) Then

Range(Cells(11, 1), Cells(25, 2)).Copy Cells(11, 7)

Else

MsgBox ("Daten in Zielzelle vorhanden. Bitte vorher löschen.")

End If

End Sub


Sub cutRangeData()

If IsEmpty(Cells(11, 7)) Then

Range(Cells(11, 1), Cells(25, 2)).Cut Cells(11, 7)

Else

MsgBox ("Daten in Zielzelle vorhanden. Bitte vorher löschen.")

End If

End Sub


Sub pasteRangeData()

If IsEmpty(Cells(11, 7)) Then

Cells(11, 7).Insert

Else

MsgBox ("Daten in Zielzelle vorhanden. Bitte vorher löschen.")

End If

End Sub


Sub createRangeData()

Dim i As Integer

For i = 11 To 25

Cells(i, 1).Value = Round(((25 - 1 + 1) * Rnd + 1), 0)
Cells(i, 2).Value = Round(((25 - 1 + 1) * Rnd + 1), 0)

Next i

End Sub


Sub removeRangeData()

Range(Cells(11, 7), Cells(25, 8)).ClearContents

End Sub
