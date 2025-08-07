Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows.Media
Imports Microsoft.Data.SqlClient
Imports System.Configuration
Class LoginView

    Private isPasswordVisible As Boolean = False
    Private txtPasswordVisible As TextBox = Nothing

    Private ReadOnly AppVersion As String = ConfigurationManager.AppSettings("AppVersion")
    Private ReadOnly Product As String = ConfigurationManager.AppSettings("Product")

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        CheckVersion()
        CheckDatabaseConnection()
        lblVersion.Text = AppVersion

        ' Initialize hidden TextBox for password visibility toggle
        If txtPasswordVisible Is Nothing Then
            txtPasswordVisible = New TextBox With {
                .FontSize = 12,
                .Foreground = New SolidColorBrush(Color.FromRgb(&H23, &H2D, &H37)),
                .Background = Brushes.Transparent,
                .BorderThickness = New Thickness(0),
                .Width = 166,
                .Padding = New Thickness(4, 0, 0, 0),
                .VerticalContentAlignment = VerticalAlignment.Center,
                .Visibility = Visibility.Collapsed
            }
            AddHandler txtPasswordVisible.TextChanged, AddressOf txtPasswordVisible_TextChanged
        End If

        ' Hooks to check Enter key to trigger login
        AddHandler txtUsername.KeyDown, AddressOf EnterKeyPressed
        AddHandler txtPassword.KeyDown, AddressOf EnterKeyPressed
        AddHandler txtPasswordVisible.KeyDown, AddressOf EnterKeyPressed
    End Sub

    Private Sub CheckVersion()
        Dim isCurrent = Login_Home_Dbhelper.CheckVersion(AppVersion, Product)
        If Not isCurrent Then
            btnLogin.IsEnabled = False
            lblUpdate.Visibility = Visibility.Visible
        Else
            btnLogin.IsEnabled = True
            lblUpdate.Visibility = Visibility.Collapsed
        End If
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

    Private Sub btnLogin_Click(sender As Object, e As RoutedEventArgs) Handles btnLogin.Click
        lblError.Visibility = Visibility.Collapsed

        If String.IsNullOrWhiteSpace(txtUsername.Text) OrElse
       (If(isPasswordVisible, String.IsNullOrWhiteSpace(txtPasswordVisible.Text), String.IsNullOrWhiteSpace(txtPassword.Password))) Then
            lblError.Text = "Please enter username and password."
            lblError.Visibility = Visibility.Visible
            Login_Home_Dbhelper.LogLogin(txtUsername.Text, False, "Empty credentials")
            Return
        End If

        Dim passwordToUse = If(isPasswordVisible, txtPasswordVisible.Text, txtPassword.Password)
        Dim user = Login_Home_Dbhelper.ValidateUser(txtUsername.Text, passwordToUse)

        If user IsNot Nothing Then
            Login_Home_Dbhelper.LogLogin(user.UserID, True, "Login success")
            MessageBox.Show($"Hi, {user.Name} ({user.Role})", "Login Successful", MessageBoxButton.OK, MessageBoxImage.Information)

            ' ✅ Pass user directly
            Dim homeWindow As New HomeView(user)
            homeWindow.Show()
            Me.Close()
        Else
            lblError.Text = "Invalid username or password."
            lblError.Visibility = Visibility.Visible
            Login_Home_Dbhelper.LogLogin(txtUsername.Text, False, "Invalid credentials")
        End If
    End Sub

    Private Sub EnterKeyPressed(sender As Object, e As KeyEventArgs)
        If e.Key = Key.Enter Then
            ' Prevent double trigger if the login button is already focused
            If Not btnLogin.IsFocused Then
                e.Handled = True
                btnLogin_Click(btnLogin, New RoutedEventArgs())
            End If
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
        Close()
        Application.Current.Shutdown()
    End Sub
End Class
