
Public Class OSPProjModel
    ' This may contain header-level info like PROJECTNO etc.
    Public Property PROJECTNO As String
    Public Property Customer As String
    Public Property Status As String

End Class





Public Class OSPModel

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



Public Class OSPLotModel

    Public Property Transaction_ID As String
    Public Property JC_no As String
    Public Property Lot_id As String
    Public Property Operation_id As String
    Public Property Quantity As Int32
    Public Property Entry_date As Date?
    Public Property Done_by As String
    Public Property Moved_to As String


End Class