Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Data.SqlClient
'Imports TACKERCALSTO.PlanningModel

Public Class dash_Plan_Dbhelper
    Private Shared ReadOnly conString As String = ConfigurationManager.ConnectionStrings("Db_Server").ConnectionString

    Public Shared Function GetWids() As List(Of PlanningModel)
        Dim list As New List(Of PlanningModel)
        Using con As New SqlConnection(conString)
            con.Open()
            Using cmd As New SqlCommand("SELECT * FROM Plan_Maindg", con)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        list.Add(New PlanningModel With {
                            .PROJECTNO = rdr("PRO_No").ToString(),
                            .Customer = rdr("Customer").ToString()
                        })
                    End While
                End Using
            End Using
        End Using
        Return list
    End Function

    Public Shared Function GetDetailsByBOM(bomNo As String) As List(Of PlanningModel.PlanningSideDgModel)
        Dim result As New List(Of PlanningModel.PlanningSideDgModel)
        Using con As New SqlConnection(conString)
            Dim cmd As New SqlCommand("SELECT * FROM v_MT_PPC WHERE [BOM No.] = @BOMNo", con)
            cmd.Parameters.AddWithValue("@BOMNo", bomNo)
            con.Open()
            Using rdr = cmd.ExecuteReader()
                While rdr.Read()
                    result.Add(New PlanningModel.PlanningSideDgModel With {
                        .WID = rdr("WID").ToString(),
                        .bomNo = rdr("BOM No.").ToString(),
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

    Public Shared Function GetJobCardsByWID(wid As String) As List(Of JobCardModel)
        Dim result As New List(Of JobCardModel)
        Using con As New SqlConnection(conString)
            Dim cmd As New SqlCommand("SELECT * FROM RoutingUploadTable WHERE WID = @WID", con)
            cmd.Parameters.AddWithValue("@WID", wid)
            con.Open()
            Using rdr = cmd.ExecuteReader()
                While rdr.Read()
                    result.Add(New JobCardModel With {
                        .JC_no = rdr("JC_no").ToString(),
                        .wid = rdr("WID").ToString(),
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
