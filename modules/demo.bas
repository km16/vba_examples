Attribute VB_Name = "demo"
Option Explicit

Sub createData()

Columns(3).Interior.ColorIndex = 0

Dim i As Integer

For i = 11 To 21

Cells(i, 3).NumberFormat = "#,##0.00 � "
Cells(i, 3).Value = ((11111 - 99999 + 1) * Rnd + 99999)

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


Sub markAvg()

Dim avgSalary As Double
Dim i As Integer

' Farbenformatierung f�r Spalte C zur�cksetzen
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

Application.Wait (Now + #12:00:01 AM#)

Next i

End Sub


Sub activateSheet()

Worksheets(ActiveSheet.Index + 1).Activate
End Sub

