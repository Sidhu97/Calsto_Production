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
        Dim user As User = Nothing

        Using conn As SqlConnection = GetConnection()
            conn.Open()

            ' Step 1: Basic user check
            Dim cmd As New SqlCommand("
            SELECT u.id AS UserID, u.FullName AS Name, u.Mail AS Email
            FROM User_Management u
            WHERE u.UserName = @UserID AND u.Pwd = @Password AND u.IsDelete = 0
        ", conn)

            cmd.Parameters.AddWithValue("@UserID", userId)
            cmd.Parameters.AddWithValue("@Password", password)

            Using rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    user = New User With {
                    .UserID = CInt(rdr("UserID")),
                    .Name = rdr("Name").ToString(),
                    .Email = rdr("Email").ToString()
                }
                Else
                    Return Nothing ' Invalid login
                End If
            End Using

            ' Step 2: Roles
            cmd = New SqlCommand("
            SELECT r.RoleName
            FROM UserRoles ur
            INNER JOIN Roles r ON ur.RoleID = r.RoleID
            WHERE ur.UserID = @UserID
        ", conn)
            cmd.Parameters.AddWithValue("@UserID", user.UserID)

            Using rdr = cmd.ExecuteReader()
                While rdr.Read()
                    user.Roles.Add(rdr("RoleName").ToString())
                End While
            End Using

            ' Step 3: Permissions
            cmd = New SqlCommand("
            SELECT rp.RoleID, rp.TabName, rp.CanView, rp.CanEdit
            FROM UserRoles ur
            INNER JOIN RolePermissions rp ON ur.RoleID = rp.RoleID
            WHERE ur.UserID = @UserID
        ", conn)
            cmd.Parameters.AddWithValue("@UserID", user.UserID)

            Using rdr = cmd.ExecuteReader()
                While rdr.Read()
                    user.Permissions.Add(New Permission With {
                    .RoleID = CInt(rdr("RoleID")),
                    .TabName = rdr("TabName").ToString(),
                    .CanView = CBool(rdr("CanView")),
                    .CanEdit = If(IsDBNull(rdr("CanEdit")), False, CBool(rdr("CanEdit")))
                })
                End While
            End Using
        End Using

        Return user
    End Function



    Public Class PermissionRepository
        Public Shared Function GetPermissionsByRole(roleId As Integer) As List(Of Permission)
            Dim perms As New List(Of Permission)
            Using conn As SqlConnection = Login_Home_Dbhelper.GetConnection()
                Dim cmd As New SqlCommand("
                SELECT RoleID, TabName, CanView, CanEdit
                FROM RolePermissions
                WHERE RoleID = @RoleID
            ", conn)
                cmd.Parameters.AddWithValue("@RoleID", roleId)

                conn.Open()
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        perms.Add(New Permission With {
                        .RoleID = CInt(rdr("RoleID")),
                        .TabName = rdr("TabName").ToString(),
                        .CanView = CBool(rdr("CanView")),
                        .CanEdit = If(IsDBNull(rdr("CanEdit")), False, CBool(rdr("CanEdit")))
                    })
                    End While
                End Using
            End Using
            Return perms
        End Function



    End Class



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
