Imports System.Configuration
Imports System.Data
Imports Microsoft.Data.SqlClient


Public Class dash_Plan_Dbhelper
    Private Shared ReadOnly conString As String = ConfigurationManager.ConnectionStrings("Db_Server").ConnectionString

#Region "Dash_Project"
    Public Shared Function GetProjectList() As List(Of DashboardModel)
        Dim result As New List(Of DashboardModel)
        Using con As New SqlConnection(conString)
            con.Open()
            Dim cmd As New SqlCommand("SELECT * FROM v_MT_OPERATIONS", con)
            Using rdr = cmd.ExecuteReader()
                While rdr.Read()
                    result.Add(New DashboardModel With {
                    .BOM_No = rdr("BOM No.").ToString(),
                    .Customer_Name = rdr("Customer Name").ToString(),
                    .Location = rdr("Location").ToString(),
                    .PO_No = rdr("PO No").ToString(),
                    .Booking_Date = If(IsDBNull(rdr("Booking Date")), Nothing, CDate(rdr("Booking Date"))),
                    .Sales = rdr("Sales").ToString(),
                    .Total_WIDs = Convert.ToInt32(rdr("Total WIDs")),
                    .Pending_WIDs = Convert.ToInt32(rdr("Pending WIDs")),
                    .PO_Dispatch_Date = If(IsDBNull(rdr("PO Dispatch Date")), Nothing, CDate(rdr("PO Dispatch Date"))),
                    .Sales_Dispatch_Date = If(IsDBNull(rdr("Sales Dispatch Date")), Nothing, CDate(rdr("Sales Dispatch Date"))),
                    .Dispatch_Date = If(IsDBNull(rdr("Dispatch Date")), Nothing, CDate(rdr("Dispatch Date"))),
                    .Project_Status = rdr("Project Status").ToString(),
                    .Due_Date = If(IsDBNull(rdr("Due Date")), Nothing, Convert.ToInt32(rdr("Due Date")))
                })
                End While
            End Using
        End Using
        Return result
    End Function
#End Region

#Region "Dash_SideDG"
    Public Shared Function GetProjectSideDetails(proNo As String) As List(Of SideDGModel)
        Dim results As New List(Of SideDGModel)
        Using con As New SqlConnection(conString)
            Dim cmd As New SqlCommand("SELECT * FROM PROJTAB_sideDG WHERE PRO_No = @proNo", con)
            cmd.Parameters.AddWithValue("@proNo", proNo)
            con.Open()
            Using rdr = cmd.ExecuteReader()
                While rdr.Read()
                    results.Add(New SideDGModel With {
                    .PRO_No = rdr("PRO_No").ToString(),
                    .Field = rdr("Field").ToString(),
                    .Detail = rdr("Detail").ToString()
                })
                End While
            End Using
        End Using
        Return results
    End Function
#End Region




#Region "Date Apply"
    Public Shared Sub UpdateProjectDispatchDate(bomNo As String, poDispatchDate As Date)
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand("usp_UpdateProjectDates", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@BOMNo", bomNo)
                cmd.Parameters.AddWithValue("@PODispatchDate", poDispatchDate)
                con.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
#End Region





End Class





