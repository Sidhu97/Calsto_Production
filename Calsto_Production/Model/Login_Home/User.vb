Public Class User
    Public Property UserID As String
    Public Property Name As String
    Public Property Email As String
    Public Property Role As String
    Public Property Permissions As List(Of Permission)

    Public Property Roles As List(Of String)



    Public Sub New()
        Permissions = New List(Of Permission)()
        Roles = New List(Of String)()
    End Sub
End Class


Public Class Permission
    Public Property RoleID As Integer
    Public Property TabName As String
    Public Property CanView As Boolean
    Public Property CanEdit As Boolean
End Class

Public Class Role
    Public Property RoleID As Integer
    Public Property RoleName As String
End Class
