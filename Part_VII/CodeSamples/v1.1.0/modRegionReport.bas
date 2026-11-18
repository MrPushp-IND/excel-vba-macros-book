Attribute VB_Name = "modRegionReport"
Option Explicit

' Columns of the Sales sheet used here.
Private Const COL_DATE As Long = 2
Private Const COL_REGION As Long = 5
Private Const COL_AMOUNT As Long = 11
Private Const COL_STATUS As Long = 12

Public Sub BuildRegionReport()
    ' Entry point: region totals for the month in Settings!B2, written to Report.
    Dim reportMonth As Date
    Dim salesData As Variant
    Dim regionNames As Variant
    Dim totals As Variant
    Dim lineCount As Long

    On Error GoTo ErrorHandler
    reportMonth = ThisWorkbook.Worksheets("Settings").Range("B2").Value
    salesData = ThisWorkbook.Worksheets("Sales").Range("A1").CurrentRegion.Value
    regionNames = Array("North", "West", "South", "East")

    totals = RegionTotals(salesData, regionNames, reportMonth, lineCount)
    WriteRegionTable ThisWorkbook.Worksheets("Report"), regionNames, totals

    MsgBox "Region report for " & Format(reportMonth, "mmmm yyyy") & " written to " & _
           "Report: " & lineCount & " order lines, Rs " & _
           Format(SumArray(totals), "#,##0") & ".", vbInformation, "Region Report"
    Exit Sub

ErrorHandler:
    MsgBox "The region report could not be built: " & Err.Description, vbExclamation, _
           "Region Report"
End Sub

Public Function RegionTotals(ByRef salesData As Variant, ByVal regionNames As Variant, _
                             ByVal reportMonth As Date, _
                             ByRef lineCount As Long) As Variant
    ' Pure calculation: non-cancelled sales by region in the month of reportMonth.
    ' Reads only its arguments and writes no cells, so it can be tested on its own.
    Dim totals() As Double
    Dim monthStart As Date
    Dim nextMonth As Date
    Dim r As Long
    Dim k As Long

    ReDim totals(LBound(regionNames) To UBound(regionNames))
    monthStart = DateSerial(Year(reportMonth), Month(reportMonth), 1)
    nextMonth = DateSerial(Year(reportMonth), Month(reportMonth) + 1, 1)
    lineCount = 0

    For r = 2 To UBound(salesData, 1)
        If salesData(r, COL_DATE) >= monthStart And salesData(r, COL_DATE) < nextMonth _
           And salesData(r, COL_STATUS) <> "Cancelled" Then
            For k = LBound(regionNames) To UBound(regionNames)
                If salesData(r, COL_REGION) = regionNames(k) Then
                    totals(k) = totals(k) + salesData(r, COL_AMOUNT)
                    lineCount = lineCount + 1
                    Exit For
                End If
            Next k
        End If
    Next r
    RegionTotals = totals
End Function

Public Sub WriteRegionTable(ByVal target As Worksheet, ByVal regionNames As Variant, _
                            ByVal totals As Variant)
    ' Output only: writes the table and formats it. Knows nothing about Sales.
    Dim grandTotal As Double
    Dim k As Long
    Dim outRow As Long

    grandTotal = SumArray(totals)
    target.Cells.Clear
    target.Range("A1:C1").Value = Array("Region", "Sales", "Share")
    outRow = 1
    For k = LBound(regionNames) To UBound(regionNames)
        outRow = outRow + 1
        target.Cells(outRow, 1).Value = regionNames(k)
        target.Cells(outRow, 2).Value = totals(k)
        If grandTotal <> 0 Then target.Cells(outRow, 3).Value = totals(k) / grandTotal
    Next k
    outRow = outRow + 1
    target.Cells(outRow, 1).Resize(1, 3).Value = Array("Total", grandTotal, 1)

    target.Range("B2").Resize(outRow - 1, 1).NumberFormat = "#,##0"
    target.Range("C2").Resize(outRow - 1, 1).NumberFormat = "0.0%"
    target.Range("A1:C1").Font.Bold = True
    target.Cells(outRow, 1).Resize(1, 3).Font.Bold = True
End Sub

Public Function SumArray(ByVal values As Variant) As Double
    Dim k As Long
    For k = LBound(values) To UBound(values)
        SumArray = SumArray + values(k)
    Next k
End Function
