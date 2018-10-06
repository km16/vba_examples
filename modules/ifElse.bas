Attribute VB_Name = "ifElse"
Option Explicit


Sub createSalaryData()

' Farbformatierung f�r Spalte C zur�cksetzen
Columns(3).Interior.ColorIndex = 0

Dim i As Integer

For i = 11 To 21

' Setze Zellenformat auf W�hrung in �
Cells(i, 3).NumberFormat = "#,##0.00 � "
' Generiere Zufallswerte
Cells(i, 3).Value = ((99999 - 1111 + 1) * Rnd + 1111)

Next i

End Sub


Sub markAvgSalary()

Dim avgSalary As Double
Dim i As Integer

' Farbformatierung f�r Spalte C zur�cksetzen
Cells(23, 3).Interior.Color = RGB(250, 230, 153)

' Durchschnittliches Gehalt aus Zelle extrahieren und in Variable speichern
avgSalary = Cells(23, 3)
' Zeile 11 bis 21 = Gehaltsdaten
For i = 11 To 21

' Setze i-Wert in Cells-Funktion ein und vergleiche den Wert darin mit Durchschnitts-
' gehalt. Wenn er gr��er oder gleich ist, f�rbe Zelle gr�n. Sonst orange.
If Cells(i, 3) >= avgSalary Then
    Cells(i, 3).Interior.Color = RGB(175, 239, 178)
Else
    Cells(i, 3).Interior.Color = RGB(248, 203, 173)
End If

' (f�r Word-Dokument streichen)
Application.Wait (Now + #12:00:01 AM#)

Next i

End Sub
