Public Class UserAccess
    Public Enum UserRole
        Admin
        Production
        Planning
        Stores
    End Enum

    Public Class RoleAccess
        Public Shared Function GetAccessibleTabs(role As UserRole) As List(Of String)
            Select Case role
                Case UserRole.Admin
                    Return New List(Of String) From {"DASHBOARD", "PLANNING", "FORMING", "WELDING", "LASER", "COATING", "OSP", "STORES", "INVENTORY", "PACKING", "DISPATCH"}
                Case UserRole.Production
                    Return New List(Of String) From {"DASHBOARD", "FORMING", "WELDING", "LASER", "COATING", "OSP", "PACKING"}
                Case UserRole.Planning
                    Return New List(Of String) From {"DASHBOARD", "PLANNING"}
                Case UserRole.Stores
                    Return New List(Of String) From {"DASHBOARD", "STORES", "INVENTORY", "DISPATCH"}
                Case Else
                    Return New List(Of String)
            End Select
        End Function
    End Class


End Class
