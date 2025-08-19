Imports System.Configuration
Imports System.Data
Imports Microsoft.Data.SqlClient

Public Class FormingDBHelper

    Private Shared ReadOnly conString As String = ConfigurationManager.ConnectionStrings("Db_Server").ConnectionString


#Region "Project List"
    Public Shared Function GetProjectNO() As List(Of FormngProjModel)
        Dim Projectlist As New List(Of FormngProjModel)
        Using con As New SqlConnection(conString)
            con.Open()
            Using cmd As New SqlCommand("SELECT * FROM V_PROJ_FORMING", con)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Projectlist.Add(New FormngProjModel With {
                            .PROJECTNO = rdr("PRO_No").ToString(),
                            .Customer = rdr("Customer").ToString(),
                            .Status = rdr("Status").ToString()
                        })
                    End While
                End Using
            End Using
        End Using
        Return Projectlist
    End Function

#End Region

#Region "Get Forming Jobs"
    Public Shared Function GetFormingJobs(proNo As String) As List(Of FormingModel)
        Dim formingJobs As New List(Of FormingModel)

        Using con As New SqlConnection(conString)
            ' Add WHERE clause to use the parameter
            Dim query As String = "SELECT * FROM V_JOB_FORMING WHERE [BOM No.] = @BOMNo"
            Dim cmd As New SqlCommand(query, con)
            cmd.Parameters.AddWithValue("@BOMNo", proNo)

            Try
                con.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read()
                    Dim job As New FormingModel With {
                        .WID = reader("WID").ToString(),
                        .BOMNo = reader("BOM No.").ToString(),
                        .BOMName = reader("BOM Name").ToString(),
                        .JC_no = reader("JC_no").ToString(),
                        .PartNumber = reader("Part Number").ToString(),
                        .Description = reader("Description").ToString(),
                        .Length = If(IsDBNull(reader("Length")), 0D, Convert.ToDecimal(reader("Length"))),
                        .Width = If(IsDBNull(reader("Width")), 0D, Convert.ToDecimal(reader("Width"))),
                        .Thick = If(IsDBNull(reader("Thick")), 0D, Convert.ToDecimal(reader("Thick"))),
                        .UnitWeight = If(IsDBNull(reader("Unit Weight")), 0D, Convert.ToDecimal(reader("Unit Weight"))),
                        .TotalWeight = If(IsDBNull(reader("Total Weight")), 0D, Convert.ToDecimal(reader("Total Weight"))),
                        .Colour = reader("Colour").ToString(),
                        .BOMQty = If(IsDBNull(reader("BOM Qty")), 0, Convert.ToInt32(reader("BOM Qty"))),
                        .ProducedQty = If(IsDBNull(reader("Produced Qty")), 0, Convert.ToInt32(reader("Produced Qty"))),
                        .BalanceQty = If(IsDBNull(reader("Balance Qty")), 0, Convert.ToInt32(reader("Balance Qty"))),
                        .Status = reader("Status").ToString()
                    }
                    formingJobs.Add(job)
                End While
            Catch ex As Exception
                MessageBox.Show("Error loading forming jobs: " & ex.Message)
            End Try
        End Using

        Return formingJobs
    End Function

#End Region


#Region "Get Forming LOT IDs"
    Public Shared Function GetFormingLots(Jobno As String) As List(Of FormingLotModel)
        Dim forminglots As New List(Of FormingLotModel)

        Using con As New SqlConnection(conString)
            ' Add WHERE clause to use the parameter
            Dim query As String = "SELECT * FROM V_JOB_Entry WHERE [JC_no] = @JC_no"
            Dim cmd As New SqlCommand(query, con)
            cmd.Parameters.AddWithValue("@JC_no", Jobno)

            Try
                con.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read()
                    Dim job As New FormingLotModel With {
                        .Transaction_ID = Convert.ToInt32(reader("Transaction_ID")),
                        .JC_no = reader("JC_no").ToString(),
                            .Lot_id = reader("Lot_id").ToString(),
                            .Operation_id = reader("Operation_ID").ToString(),
                            .Quantity = If(IsDBNull(reader("Quantity")), 0, Convert.ToInt32(reader("Quantity"))),
                            .Entry_date = If(IsDBNull(reader("Entry_Date")), Nothing, Convert.ToDateTime(reader("Entry_Date"))),
                            .Done_by = If(IsDBNull(reader("Done_by")), String.Empty, reader("Done_by").ToString()),
                            .Moved_to = If(IsDBNull(reader("Moved_to")), String.Empty, reader("Moved_to").ToString())
                    }
                    forminglots.Add(job)
                End While
            Catch ex As Exception
                MessageBox.Show("Error loading Lot IDs: " & ex.Message)
            End Try
        End Using

        Return forminglots
    End Function

#End Region


    Public Shared Sub CreateJobEntry(JC As String, Qty As Int32, createdBy As String, remark As String)
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand("sp_CreateLotEntry", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@JC_no", JC)
                cmd.Parameters.AddWithValue("@Quantity", Qty)
                cmd.Parameters.AddWithValue("@Entered_By", createdBy)
                cmd.Parameters.AddWithValue("@Remarks", remark)
                con.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub




#Region "Edit Forming LOT"
    Public Shared Function EditLotEntry(TransId As String, lotId As String, newQty As Integer, updatedBy As String, remarks As String) As Boolean
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand("sp_UpdateJobCardEntry", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@Transaction_ID", TransId)
                cmd.Parameters.AddWithValue("@JC_no", lotId)
                cmd.Parameters.AddWithValue("@Quantity", newQty)
                cmd.Parameters.AddWithValue("@Remarks", remarks)
                cmd.Parameters.AddWithValue("@UpdatedBy", updatedBy)

                Try
                    con.Open()
                    cmd.ExecuteNonQuery()
                    Return True
                Catch ex As Exception
                    MessageBox.Show("Error updating Lot Entry: " & ex.Message)
                    Return False
                End Try
            End Using
        End Using
    End Function
#End Region







End Class
