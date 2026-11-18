Attribute VB_Name = "modLegacy"
Sub RegionReport()
' region report
    Sheets("Sales").Select
    Range("A1").Select
    n = Range("A1").End(xlDown).Row
    For i = 2 To n
        If Month(Cells(i, 2)) = 10 And Cells(i, 12) <> "Cancelled" Then
            If Cells(i, 5) = "North" Then t1 = t1 + Cells(i, 11)
            If Cells(i, 5) = "West" Then t2 = t2 + Cells(i, 11)
            If Cells(i, 5) = "South" Then t3 = t3 + Cells(i, 11)
            If Cells(i, 5) = "East" Then t4 = t4 + Cells(i, 11)
            c = c + 1
        End If
    Next
    Sheets("Report").Select
    Cells.Clear
    Range("A1") = "Region"
    Range("B1") = "Sales"
    Range("C1") = "Share"
    Range("A2") = "North"
    Range("B2") = t1
    Range("C2") = t1 / (t1 + t2 + t3 + t4)
    Range("A3") = "West"
    Range("B3") = t2
    Range("C3") = t2 / (t1 + t2 + t3 + t4)
    Range("A4") = "South"
    Range("B4") = t3
    Range("C4") = t3 / (t1 + t2 + t3 + t4)
    Range("A5") = "East"
    Range("B5") = t4
    Range("C5") = t4 / (t1 + t2 + t3 + t4)
    Range("A6") = "Total"
    Range("B6") = t1 + t2 + t3 + t4
    Range("C6") = 1
    Range("B2:B6").NumberFormat = "#,##0"
    Range("C2:C6").NumberFormat = "0.0%"
    Range("A1:C1").Font.Bold = True
    Range("A6:C6").Font.Bold = True
    MsgBox "Done " & c
End Sub
