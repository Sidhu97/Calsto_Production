Public Class DashboardModel
    Public Property BOM_No As String          ' [BOM No.]
    Public Property Customer_Name As String   ' [Customer Name]
    Public Property Location As String
    Public Property PO_No As String
    Public Property Booking_Date As Date?
    Public Property Sales As String

    Public Property Total_WIDs As Integer
    Public Property Pending_WIDs As Integer

    Public Property PO_Dispatch_Date As Date?
    Public Property Sales_Dispatch_Date As Date?
    Public Property Dispatch_Date As Date?

    Public Property Project_Status As String
    Public Property Due_Date As Integer?
End Class
Public Class SideDGModel
    Public Property PRO_No As String
    Public Property Field As String
    Public Property Detail As String
End Class
