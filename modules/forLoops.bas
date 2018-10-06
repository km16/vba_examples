Attribute VB_Name = "forLoops"
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


Sub sumSalary()

Dim i As Integer
Dim sum As Double

' Deklariere Summenvariable
sum = 0

' Loop durch jede Zelle in der Gehaltsspalte
For i = 11 To 21

' Addiere Wert in der derzeit gepr�ften Zelle zu Variable sum
sum = sum + Cells(i, 3)

' Hebe derzeit gepr�fte Zelle hervor
'(f�r Word-Dokument streichen)
Cells(i, 3).Interior.Color = RGB(255, 230, 153)

' Zeige neue Summe nach jedem Loop in der Summe-Zelle an
Cells(23, 3).Value = sum

' Warte eine Sekunde (Ewig lange Loops m�glich, mit Bedacht experimentieren)
'(f�r Word-Dokument streichen)
Application.Wait (Now + #12:00:01 AM#)

' Setze Hintergrund der Zelle zur�ck
'(f�r Word-Dokument streichen)
Cells(i, 3).Interior.ColorIndex = 0

Next i

End Sub

