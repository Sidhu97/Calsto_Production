' Packing slip header model
Public Class PackingslipHeaderModel
    Public Property PackID As String
    Public Property ProjectNo As String
    Public Property Customer As String
    Public Property Location As String
    Public Property SystemType As String
    Public Property PackType As String
    Public Property PackedBy As String
    Public Property PackDate As DateTime
    Public Property Items As List(Of PackingItemModel)
End Class

' Packing slip item model
Public Class PackingItemModel
    Public Property WID As String
    Public Property PartNumber As String
    Public Property Description As String
    Public Property Qty As Integer
End Class
