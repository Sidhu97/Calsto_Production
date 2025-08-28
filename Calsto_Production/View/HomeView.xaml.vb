Class HomeView

    Private _currentUser As User

    ' Constructor accepts logged-in user
    Public Sub New(user As User)
        InitializeComponent()
        _currentUser = user

        ApplyRoleBasedAccess()
    End Sub

    Private Sub ApplyRoleBasedAccess()
        ' Hide all tabs first
        TabDashboard.Visibility = Visibility.Collapsed
        TabPlanning.Visibility = Visibility.Collapsed
        TabForming.Visibility = Visibility.Collapsed
        TabWelding.Visibility = Visibility.Collapsed
        TabLaser.Visibility = Visibility.Collapsed
        TabCoating.Visibility = Visibility.Collapsed
        TabOsp.Visibility = Visibility.Collapsed
        TabStores.Visibility = Visibility.Collapsed
        TabInventory.Visibility = Visibility.Collapsed
        TabPacking.Visibility = Visibility.Collapsed
        TabDispatch.Visibility = Visibility.Collapsed

        ' Show tabs based on user role
        Select Case _currentUser.Role
            Case "Admin"
                ' Admin sees everything
                TabDashboard.Visibility = Visibility.Visible
                TabPlanning.Visibility = Visibility.Visible
                TabForming.Visibility = Visibility.Visible
                TabWelding.Visibility = Visibility.Visible
                TabLaser.Visibility = Visibility.Visible
                TabCoating.Visibility = Visibility.Visible
                TabOsp.Visibility = Visibility.Visible
                TabStores.Visibility = Visibility.Visible
                TabInventory.Visibility = Visibility.Visible
                TabPacking.Visibility = Visibility.Visible
                TabDispatch.Visibility = Visibility.Visible

            Case "Manager"
                TabDashboard.Visibility = Visibility.Visible
                TabPlanning.Visibility = Visibility.Visible
                TabForming.Visibility = Visibility.Visible
                TabWelding.Visibility = Visibility.Visible
                TabCoating.Visibility = Visibility.Visible
                TabInventory.Visibility = Visibility.Visible
                TabPacking.Visibility = Visibility.Visible

            Case "User"
                TabDashboard.Visibility = Visibility.Visible
                TabPlanning.Visibility = Visibility.Visible
                TabInventory.Visibility = Visibility.Visible
                TabPacking.Visibility = Visibility.Visible
        End Select
    End Sub

End Class
