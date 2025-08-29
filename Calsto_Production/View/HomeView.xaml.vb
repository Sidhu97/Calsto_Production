Imports System.Windows
Imports System.Windows.Controls

Public Class HomeView
    Private ReadOnly _currentUser As User

    ' Constructor accepts logged-in user
    Public Sub New(user As User)
        InitializeComponent()
        _currentUser = user
        loggeduser.Text = user.Name
        ApplyRoleBasedAccess()
    End Sub

    ''' <summary>
    ''' Apply dynamic access control based on DB permissions
    ''' </summary>
    Private Sub ApplyRoleBasedAccess()
        ' Hide all tabs first
        For Each item As TabItem In HomeTabControl.Items
            item.Visibility = Visibility.Collapsed
        Next

        ' Show tabs if user has CanView permission
        For Each perm As Permission In _currentUser.Permissions
            Dim tab As TabItem = TryCast(HomeTabControl.FindName("Tab" & perm.TabName), TabItem)

            If tab IsNot Nothing AndAlso perm.CanView Then
                tab.Visibility = Visibility.Visible

                ' Optional: disable edit-only sections inside the tab
                If Not perm.CanEdit Then
                    Dim editPanel As FrameworkElement = TryCast(tab.FindName("EditPanel"), FrameworkElement)
                    If editPanel IsNot Nothing Then
                        editPanel.IsEnabled = False
                    End If
                End If
            End If
        Next
    End Sub

End Class
