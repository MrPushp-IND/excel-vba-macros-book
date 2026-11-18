VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mOrderId As String
Private mOrderDate As Date
Private mCustomer As String
Private mRep As String
Private mQuantity As Long
Private mUnitPrice As Double
Private mStatus As String

Private Const ERR_INVALID As Long = vbObjectError + 513

' ---------- Properties ----------
Public Property Get OrderId() As String: OrderId = mOrderId: End Property
Public Property Let OrderId(ByVal value As String)
    If Len(Trim$(value)) = 0 Then Err.Raise ERR_INVALID, "clsOrder", "Order ID is empty."
    mOrderId = Trim$(value)
End Property

Public Property Get OrderDate() As Date: OrderDate = mOrderDate: End Property
Public Property Let OrderDate(ByVal value As Date): mOrderDate = value: End Property

Public Property Get Customer() As String: Customer = mCustomer: End Property
Public Property Let Customer(ByVal value As String): mCustomer = value: End Property

Public Property Get Rep() As String: Rep = mRep: End Property
Public Property Let Rep(ByVal value As String): mRep = value: End Property

Public Property Get Quantity() As Long: Quantity = mQuantity: End Property
Public Property Let Quantity(ByVal value As Long)
    If value <= 0 Then Err.Raise ERR_INVALID, "clsOrder", _
        "Quantity must be greater than zero (order " & mOrderId & ")."
    mQuantity = value
End Property

Public Property Get UnitPrice() As Double: UnitPrice = mUnitPrice: End Property
Public Property Let UnitPrice(ByVal value As Double)
    If value < 0 Then Err.Raise ERR_INVALID, "clsOrder", _
        "Unit price cannot be negative (order " & mOrderId & ")."
    mUnitPrice = value
End Property

Public Property Get Status() As String: Status = mStatus: End Property
Public Property Let Status(ByVal value As String)
    Select Case value
        Case "Pending", "Delivered", "Cancelled"
            mStatus = value
        Case Else
            Err.Raise ERR_INVALID, "clsOrder", "Unknown status '" & value & "'."
    End Select
End Property

Public Property Get Amount() As Double            ' Read-only: calculated
    Amount = mQuantity * mUnitPrice
End Property

' ---------- Methods ----------
Public Sub LoadFromRow(ByRef data As Variant, ByVal rowIndex As Long)
    ' Fills the order from one row of the Sales array (columns A to L).
    OrderId = data(rowIndex, 1)
    OrderDate = data(rowIndex, 2)
    Customer = data(rowIndex, 4)
    Rep = data(rowIndex, 6)
    Quantity = data(rowIndex, 9)
    UnitPrice = data(rowIndex, 10)
    Status = data(rowIndex, 12)
End Sub

Public Function DaysOpen(ByVal asOf As Date) As Long
    DaysOpen = asOf - mOrderDate
End Function

Public Function IsOverdue(ByVal asOf As Date, ByVal limitDays As Long) As Boolean
    ' The business rule, in one place: pending for more than limitDays days.
    IsOverdue = (mStatus = "Pending" And DaysOpen(asOf) > limitDays)
End Function

Public Function Describe() As String
    Describe = mOrderId & ", " & Format(mOrderDate, "dd-mmm-yyyy") & ", " & mCustomer & _
               ", Rs " & Format(Amount, "#,##0") & ", " & mStatus
End Function
