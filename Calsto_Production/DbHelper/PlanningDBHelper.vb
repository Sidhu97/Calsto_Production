Imports System.Configuration
Imports System.Data
Imports Microsoft.Data.SqlClient

Public Class PlanningDBHelper

    Private Shared ReadOnly conString As String = ConfigurationManager.ConnectionStrings("Db_Server").ConnectionString


#Region "Project List"
    Public Shared Function GetProjectNO() As List(Of PlanningModel)
        Dim Projectlist As New List(Of PlanningModel)
        Using con As New SqlConnection(conString)
            con.Open()
            Using cmd As New SqlCommand("SELECT * FROM PRO_list", con)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Projectlist.Add(New PlanningModel With {
                            .PROJECTNO = rdr("PRO_No").ToString(),
                            .Customer = rdr("Customer").ToString()
                        })
                    End While
                End Using
            End Using
        End Using
        Return Projectlist
    End Function

#End Region


#Region "BOM Items"
    Public Shared Function GetBOMItems(bomNo As String) As List(Of PlanningSideDgModel)
        Dim result As New List(Of PlanningSideDgModel)
        Using con As New SqlConnection(conString)
            Dim cmd As New SqlCommand("SELECT * FROM v_MT_PPC WHERE [BOM No.] = @BOMNo", con)
            cmd.Parameters.AddWithValue("@BOMNo", bomNo)
            con.Open()
            Using rdr = cmd.ExecuteReader()
                While rdr.Read()
                    result.Add(New PlanningSideDgModel With {
                        .WID = rdr("WID").ToString(),
                        .BOMNo = rdr("BOM No.").ToString(),
                        .BOMName = rdr("BOM Name").ToString(),
                        .Customer = rdr("Customer").ToString(),
                        .Description = rdr("Description").ToString(),
                        .Assigned = rdr("Assinged").ToString(),
                        .BOMQty = Convert.ToInt32(rdr("BOM Qty")),
                        .Colour = rdr("Colour").ToString(),
                        .Process = rdr("Process").ToString(),
                        .Jobcard = rdr("JobCard").ToString(),
                        .PackingBalance = Convert.ToInt32(rdr("Packing Balance.")),
                        .DispatchQty = Convert.ToInt32(rdr("Dispatched Qty")),
                        .DispatchPending = Convert.ToInt32(rdr("Dispatch Pending.")),
                        .DispatchDate = If(IsDBNull(rdr("Dispatch Date")), Nothing, CDate(rdr("Dispatch Date")))
                    })
                End While
            End Using
        End Using
        Return result
    End Function
#End Region


    Public Shared Function GetJobCards(wid As String) As List(Of JobCardModel)
        Dim result As New List(Of JobCardModel)
        Using con As New SqlConnection(conString)
            Dim cmd As New SqlCommand("SELECT * FROM [JobCard_log] WHERE WID = @WID", con)
            cmd.Parameters.AddWithValue("@WID", wid)
            con.Open()
            Using rdr = cmd.ExecuteReader()
                While rdr.Read()
                    result.Add(New JobCardModel With {
                        .JC_no = rdr("JC_no").ToString(),
                        .WID = rdr("WID").ToString(),
                        .OP_Sequence = rdr("OP_Sequence").ToString(),
                        .Operation_ID = rdr("Operation_ID").ToString(),
                        .Created_date = rdr("Created_date"),
                        .Created_by = rdr("Created_by").ToString()
                    })
                End While
            End Using
        End Using
        Return result
    End Function

    Public Shared Sub CreateJobCard(wid As String, createdBy As String)
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand("sp_CreateJobCard_ByWID", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@WID", wid)
                cmd.Parameters.AddWithValue("@CreatedBy", createdBy)
                con.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
End Class