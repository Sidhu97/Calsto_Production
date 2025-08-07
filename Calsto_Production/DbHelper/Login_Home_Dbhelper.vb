Imports System.Data
Imports System.Configuration
Imports Microsoft.Data.SqlClient

Public Class Login_Home_Dbhelper

    Private Shared ReadOnly connectionString As String = ConfigurationManager.ConnectionStrings("Db_Server").ConnectionString

    Public Shared Function GetConnection() As SqlConnection
        Return New SqlConnection(connectionString)
    End Function

    ' Validate user credentials; returns User object if successful, Nothing otherwise.
    Public Shared Function ValidateUser(userId As String, password As String) As User
        Using conn As SqlConnection = GetConnection()
            Dim cmd As New SqlCommand("SELECT UserID, Name, Email, Role FROM USERDETAILS WHERE UserID = @UserID AND Password = @Password", conn)
            cmd.Parameters.AddWithValue("@UserID", userId)
            cmd.Parameters.AddWithValue("@Password", password)
            conn.Open()
            Using rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    ' Success: return User object
                    Return New User With {
                        .UserID = rdr("UserID").ToString(),
                        .Name = rdr("Name").ToString(),
                        .Email = rdr("Email").ToString(),
                        .Role = rdr("Role").ToString()
                    }
                End If
            End Using
        End Using
        ' Failed
        Return Nothing
    End Function

    ' Optional: Logs every login attempt
    Public Shared Sub LogLogin(userId As String, success As Boolean, message As String)
        Using conn As SqlConnection = GetConnection()
            Dim cmd As New SqlCommand("INSERT INTO LOGIN_LOG (USERID, SYSTEMNAME, MACADDRESS, STATUS, MESSAGE) VALUES (@USERID, @SYSTEMNAME, @MAC, @STATUS, @MSG)", conn)
            cmd.Parameters.AddWithValue("@USERID", userId)
            cmd.Parameters.AddWithValue("@SYSTEMNAME", Environment.MachineName)
            cmd.Parameters.AddWithValue("@MAC", GetMacAddress())
            cmd.Parameters.AddWithValue("@STATUS", If(success, "Success", "Failed"))
            cmd.Parameters.AddWithValue("@MSG", message)
            conn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' Version check: Does the specified product match the required version?
    Public Shared Function CheckVersion(appVersion As String, product As String) As Boolean
        Using conn As SqlConnection = GetConnection()
            Dim cmd As New SqlCommand("SELECT [Version] FROM VERSION_TABLE WHERE Product = @Product", conn)
            cmd.Parameters.AddWithValue("@Product", product)
            conn.Open()
            Dim dbVersion = TryCast(cmd.ExecuteScalar(), String)
            Return dbVersion = appVersion
        End Using
    End Function

    ' Get user details (whole row as datatable - for HomeView, if needed)
    Public Shared Function GetUserDetails(userId As String) As DataTable
        Dim table As New DataTable()
        Using conn As SqlConnection = GetConnection()
            Dim cmd As New SqlCommand("SELECT * FROM USERDETAILS WHERE UserID = @UserID", conn)
            cmd.Parameters.AddWithValue("@UserID", userId)
            conn.Open()
            Using reader As SqlDataReader = cmd.ExecuteReader()
                table.Load(reader)
            End Using
        End Using
        Return table
    End Function

    ' Get first active MAC address
    Public Shared Function GetMacAddress() As String
        For Each nic In Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
            If nic.OperationalStatus = Net.NetworkInformation.OperationalStatus.Up AndAlso
                Not String.IsNullOrEmpty(nic.GetPhysicalAddress().ToString()) Then
                Return nic.GetPhysicalAddress().ToString()
            End If
        Next
        Return ""
    End Function

End Class
