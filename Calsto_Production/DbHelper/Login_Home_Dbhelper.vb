Imports System.Data
Imports System.Configuration
Imports Microsoft.Data.SqlClient
Imports System.Net.NetworkInformation

' Define a simple User class

Public Class Login_Home_Dbhelper

    Private Shared ReadOnly connectionString As String = ConfigurationManager.ConnectionStrings("Db_Server").ConnectionString

    ' Get a new SQL connection
    Public Shared Function GetConnection() As SqlConnection
        Return New SqlConnection(connectionString)
    End Function

    ' Validate user credentials; returns User object if successful, Nothing otherwise
    Public Shared Function ValidateUserWithRole(userId As String, password As String) As User
        Using conn As SqlConnection = GetConnection()
            Dim cmd As New SqlCommand("
                SELECT u.id AS UserID, u.FullName AS Name, u.Mail AS Email, r.RoleName AS Role
                FROM User_Management u
                INNER JOIN UserRoles ur ON u.id = ur.UserID
                INNER JOIN Roles r ON ur.RoleID = r.RoleID
                WHERE u.UserName = @UserID AND u.Pwd = @Password AND u.IsDelete = 0
            ", conn)

            cmd.Parameters.AddWithValue("@UserID", userId)
            cmd.Parameters.AddWithValue("@Password", password)

            conn.Open()
            Using rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    ' Return a User object with role
                    Return New User With {
                        .UserID = rdr("UserID").ToString(),
                        .Name = rdr("Name").ToString(),
                        .Email = rdr("Email").ToString(),
                        .Role = rdr("Role").ToString()
                    }
                End If
            End Using
        End Using

        ' Login failed
        Return Nothing
    End Function

    ' Log every login attempt
    Public Shared Sub LogLogin(userId As String, success As Boolean, message As String)
        Using conn As SqlConnection = GetConnection()
            Dim cmd As New SqlCommand("
                INSERT INTO LOGIN_LOG (USERID, SYSTEMNAME, MACADDRESS, STATUS, MESSAGE)
                VALUES (@USERID, @SYSTEMNAME, @MAC, @STATUS, @MSG)
            ", conn)

            cmd.Parameters.AddWithValue("@USERID", userId)
            cmd.Parameters.AddWithValue("@SYSTEMNAME", Environment.MachineName)
            cmd.Parameters.AddWithValue("@MAC", GetMacAddress())
            cmd.Parameters.AddWithValue("@STATUS", If(success, "Success", "Failed"))
            cmd.Parameters.AddWithValue("@MSG", message)

            conn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' Get the first active MAC address
    Public Shared Function GetMacAddress() As String
        For Each nic As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()
            If nic.OperationalStatus = OperationalStatus.Up AndAlso
               Not String.IsNullOrEmpty(nic.GetPhysicalAddress().ToString()) Then
                Return nic.GetPhysicalAddress().ToString()
            End If
        Next
        Return ""
    End Function

End Class
