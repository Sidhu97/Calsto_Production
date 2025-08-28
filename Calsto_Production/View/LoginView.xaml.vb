Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports Microsoft.Data.SqlClient
Imports System.Configuration

Class LoginView

    Private isPasswordVisible As Boolean = False
    Private txtPasswordVisible As TextBox = Nothing

    ' App version and product (optional for version check)
    Private ReadOnly AppVersion As String = ConfigurationManager.AppSettings("AppVersion")
    Private ReadOnly Product As String = ConfigurationManager.AppSettings("Product")

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ' Initialize hidden TextBox for password toggle
        If txtPasswordVisible Is Nothing Then
            txtPasswordVisible = New TextBox With {
                .FontSize = 12,
                .Foreground = New SolidColorBrush(Color.FromRgb(&H23, &H2D, &H37)),
                .Background = Brushes.Transparent,
                .BorderThickness = New Thickness(0),
                .Width = txtPassword.Width,
                .Padding = New Thickness(4, 0, 0, 0),
                .VerticalContentAlignment = VerticalAlignment.Center,
                .Visibility = Visibility.Collapsed
            }
            AddHandler txtPasswordVisible.TextChanged, AddressOf txtPasswordVisible_TextChanged
        End If

        ' Enter key triggers login
        AddHandler txtUsername.KeyDown, AddressOf EnterKeyPressed
        AddHandler txtPassword.KeyDown, AddressOf EnterKeyPressed
        AddHandler txtPasswordVisible.KeyDown, AddressOf EnterKeyPressed

        ' Optional: check DB connection on load
        CheckDatabaseConnection()
    End Sub

    Private Sub CheckDatabaseConnection()
        Try
            Using con = Login_Home_Dbhelper.GetConnection()
                con.Open()
                dbStatusIndicator.Fill = New SolidColorBrush(Colors.Green)
                dbStatusText.Text = "Connected"
            End Using
        Catch
            dbStatusIndicator.Fill = New SolidColorBrush(Colors.Red)
            dbStatusText.Text = "Not Connected"
        End Try
    End Sub

    Private isLoggingIn As Boolean = False

    Private Sub btnLogin_Click(sender As Object, e As RoutedEventArgs) Handles btnLogin.Click
        If isLoggingIn Then Return
        isLoggingIn = True

        lblError.Visibility = Visibility.Hidden

        Dim passwordToUse = If(isPasswordVisible, txtPasswordVisible.Text, txtPassword.Password)
        Dim user = Login_Home_Dbhelper.ValidateUserWithRole(txtUsername.Text, passwordToUse)

        If user IsNot Nothing Then
            Login_Home_Dbhelper.LogLogin(user.UserID, True, "Login success")
            Dim homeWindow As New HomeView(user)
            homeWindow.Show()
            Me.Close()
        Else
            lblError.Text = "Invalid username or password."
            lblError.Visibility = Visibility.Visible
            Login_Home_Dbhelper.LogLogin(txtUsername.Text, False, "Invalid credentials")
            isLoggingIn = False
        End If
    End Sub

    Private Sub EnterKeyPressed(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter AndAlso Not isLoggingIn Then
            e.Handled = True
            btnLogin_Click(btnLogin, New RoutedEventArgs())
        End If
    End Sub


    Private Sub btnTogglePassword_Click(sender As Object, e As RoutedEventArgs) Handles btnTogglePassword.Click
        isPasswordVisible = Not isPasswordVisible
        Dim dockPanel = TryCast(txtPassword.Parent, DockPanel)
        If dockPanel Is Nothing Then Return

        If isPasswordVisible Then
            txtPasswordVisible.Text = txtPassword.Password
            txtPassword.Visibility = Visibility.Collapsed

            If Not dockPanel.Children.Contains(txtPasswordVisible) Then
                dockPanel.Children.Insert(2, txtPasswordVisible)
            End If
            txtPasswordVisible.Visibility = Visibility.Visible
            btnTogglePassword.Content = "🙈"
        Else
            txtPassword.Password = txtPasswordVisible.Text
            txtPassword.Visibility = Visibility.Visible
            txtPasswordVisible.Visibility = Visibility.Collapsed
            btnTogglePassword.Content = "👁️"
        End If
    End Sub

    Private Sub txtPasswordVisible_TextChanged(sender As Object, e As TextChangedEventArgs)
        If Not isPasswordVisible Then
            txtPassword.Password = txtPasswordVisible.Text
        End If
    End Sub

    Private Sub txtPassword_PreviewKeyDown(sender As Object, e As KeyEventArgs) Handles txtPassword.PreviewKeyDown
        lblCapsLock.Visibility = If(Console.CapsLock, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub txtPassword_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtPassword.GotFocus
        lblCapsLock.Visibility = If(Console.CapsLock, Visibility.Visible, Visibility.Collapsed)
    End Sub

    Private Sub txtForgotPassword_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs)
        MessageBox.Show("Please contact admin.", "Forgot Password", MessageBoxButton.OK, MessageBoxImage.Information)
    End Sub

    Private Sub label_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles label.MouseDown
        Me.Close()
        Application.Current.Shutdown()
    End Sub

End Class
