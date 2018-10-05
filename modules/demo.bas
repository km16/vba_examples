Attribute VB_Name = "demo"
Option Explicit

Sub createData()

Columns(3).Interior.ColorIndex = 0

Dim i As Integer

For i = 11 To 21

Cells(i, 3).NumberFormat = "#,##0.00 € "
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

' Addiere Wert in der derzeit geprüften Zelle zu Variable sum
sum = sum + Cells(i, 3)

' Hebe derzeit geprüfte Zelle hervor
'(für Word-Dokument streichen)
Cells(i, 3).Interior.Color = RGB(255, 230, 153)

' Zeige neue Summe nach jedem Loop in der Summe-Zelle an
Cells(23, 3).Value = sum

' Warte eine Sekunde (Ewig lange Loops möglich, mit Bedacht experimentieren)
'(für Word-Dokument streichen)
Application.Wait (Now + #12:00:01 AM#)

' Setze Hintergrund der Zelle zurück
'(für Word-Dokument streichen)
Cells(i, 3).Interior.ColorIndex = 0

Next i

End Sub


Sub markAvg()

Dim avgSalary As Double
Dim i As Integer

' Farbenformatierung für Spalte C zurücksetzen
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

Application.Wait (Now + #12:00:01 AM#)

Next i

End Sub


Sub activateSheet()

Worksheets(ActiveSheet.Index + 1).Activate
End Sub

