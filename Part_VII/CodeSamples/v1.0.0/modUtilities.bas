Attribute VB_Name = "modUtilities"
Option Explicit

' General helpers with no business knowledge.

Public Sub QuickSort2D(ByRef data As Variant, ByVal sortColumn As Long, _
                       Optional ByVal descending As Boolean = False, _
                       Optional ByVal firstRow As Long = -1, _
                       Optional ByVal lastRow As Long = -1)
    ' Sorts the rows of a 2-D array in place by one column.
    Dim low As Long
    Dim high As Long
    Dim pivotValue As Variant
    Dim column As Long
    Dim swapValue As Variant

    If firstRow = -1 Then firstRow = LBound(data, 1)
    If lastRow = -1 Then lastRow = UBound(data, 1)
    If firstRow >= lastRow Then Exit Sub

    low = firstRow
    high = lastRow
    pivotValue = data((firstRow + lastRow) \ 2, sortColumn)

    Do While low <= high
        If descending Then
            Do While data(low, sortColumn) > pivotValue: low = low + 1: Loop
            Do While data(high, sortColumn) < pivotValue: high = high - 1: Loop
        Else
            Do While data(low, sortColumn) < pivotValue: low = low + 1: Loop
            Do While data(high, sortColumn) > pivotValue: high = high - 1: Loop
        End If
        If low <= high Then
            For column = LBound(data, 2) To UBound(data, 2)       ' Swap whole rows
                swapValue = data(low, column)
                data(low, column) = data(high, column)
                data(high, column) = swapValue
            Next column
            low = low + 1
            high = high - 1
        End If
    Loop

    If firstRow < high Then QuickSort2D data, sortColumn, descending, firstRow, high
    If low < lastRow Then QuickSort2D data, sortColumn, descending, low, lastRow
End Sub
