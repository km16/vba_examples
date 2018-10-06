Attribute VB_Name = "ifElse"
Option Explicit


Sub createSalaryData()

' Farbformatierung für Spalte C zurücksetzen
Columns(3).Interior.ColorIndex = 0

Dim i As Integer

For i = 11 To 21

' Setze Zellenformat auf Währung in €
Cells(i, 3).NumberFormat = "#,##0.00 € "
' Generiere Zufallswerte
Cells(i, 3).Value = ((99999 - 1111 + 1) * Rnd + 1111)

Next i

End Sub


Sub markAvgSalary()

Dim avgSalary As Double
Dim i As Integer

' Farbformatierung für Spalte C zurücksetzen
Cells(23, 3).Interior.Color = RGB(250, 230, 153)

' Durchschnittliches Gehalt aus Zelle extrahieren und in Variable speichern
avgSalary = Cells(23, 3)
' Zeile 11 bis 21 = Gehaltsdaten
For i = 11 To 21

' Setze i-Wert in Cells-Funktion ein und vergleiche den Wert darin mit Durchschnitts-
' gehalt. Wenn er größer oder gleich ist, färbe Zelle grün. Sonst orange.
If Cells(i, 3) >= avgSalary Then
    Cells(i, 3).Interior.Color = RGB(175, 239, 178)
Else
    Cells(i, 3).Interior.Color = RGB(248, 203, 173)
End If

' (für Word-Dokument streichen)
Application.Wait (Now + #12:00:01 AM#)

Next i

End Sub
