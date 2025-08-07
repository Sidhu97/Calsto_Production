Public Class HomeView

    Private ReadOnly _currentUser As User

    Public Sub New(ByVal user As User)
        InitializeComponent()

        _currentUser = user

        ' Show "Name (Role)" only for Admins
        If user.Role.ToLower() = "admin" Then
            loggeduser.Text = user.Name & " (" & user.Role & ")"
        Else
            loggeduser.Text = user.Name
        End If
    End Sub

End Class
