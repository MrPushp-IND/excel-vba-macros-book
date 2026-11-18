Attribute VB_Name = "modDataAccess"
Option Explicit
Option Private Module

Public Enum SalesColumn
    scOrderId = 1
    scOrderDate
    scCustomerId
    scCustomer
    scRegion
    scRep
    scProduct
    scCategory
    scQuantity
    scUnitPrice
    scAmount
    scStatus
End Enum

Public Const ERR_BAD_SHAPE As Long = vbObjectError + 6601

Public Function SalesHeadings() As Variant
    ' The headings the code expects, in SalesColumn order.
    SalesHeadings = Array("Order ID", "Order Date", "Customer ID", "Customer", _
                          "Region", "Rep", "Product", "Category", "Quantity", _
                          "Unit Price", "Amount", "Status")
End Function

Public Function ReadSheetTable(ByVal sheetName As String, _
                               ByVal expected As Variant) As Variant
    ' Returns the sheet's table (headings in row 1) after checking its headings.
    Dim data As Variant
    Dim k As Long
    Dim column As Long

    data = ThisWorkbook.Worksheets(sheetName).Range("A1").CurrentRegion.Value
    For k = LBound(expected) To UBound(expected)
        column = k - LBound(expected) + 1
        If column > UBound(data, 2) Then
            Err.Raise ERR_BAD_SHAPE, "ReadSheetTable", _
                      sheetName & " has no column " & column & " ('" & expected(k) & "')."
        ElseIf StrComp(Trim$(CStr(data(1, column))), expected(k), vbTextCompare) <> 0 Then
            Err.Raise ERR_BAD_SHAPE, "ReadSheetTable", sheetName & " column " & column & _
                      " is '" & data(1, column) & "', but the code expects '" & _
                      expected(k) & "'."
        End If
    Next k
    ReadSheetTable = data
End Function

Public Function ReadSetting(ByVal label As String) As Variant
    ' Finds a setting by its label in column A of Settings and returns column B.
    Dim found As Range
    Set found = ThisWorkbook.Worksheets("Settings").Columns("A").Find(What:=label, _
                LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If found Is Nothing Then Err.Raise ERR_BAD_SHAPE, "ReadSetting", _
                                       "Setting '" & label & "' is missing from Settings."
    ReadSetting = found.Offset(0, 1).Value
End Function
