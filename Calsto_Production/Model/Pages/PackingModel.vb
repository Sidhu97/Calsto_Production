Public Class PackingModel
    Public Property WID As String
    Public Property BOMNo As String
    Public Property Customer As String
    Public Property BOMName As String
    Public Property JC_no As String
    Public Property PartNumber As String
    Public Property Description As String
    Public Property Length As Decimal
    Public Property Width As Decimal
    Public Property Thick As Decimal
    Public Property UnitWeight As Decimal
    Public Property TotalWeight As Decimal
    Public Property Colour As String
    Public Property BOMQty As Integer
    Public Property ProducedQty As Integer
    Public Property BalanceQty As Integer
    Public Property Status As String

End Class



Public Class PackingProjModel
    Public Property Proj_no As String
    Public Property Customer As String

    Public Property Status As String

End Class

Public Class PackingTypeModel
    Public Property Packtype As String

End Class





Public Class PackingHeaderModel
    Public Property PackID As String
    Public Property PackEntryDate As Date?
    Public Property PackType As String
    Public Property ProjNo As String
    Public Property CreatedBy As String
End Class
