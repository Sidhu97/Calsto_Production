Imports System.Configuration
Imports System.Data
Imports Calsto_Production.PackingModel
Imports DocumentFormat.OpenXml.Drawing.Diagrams
Imports Microsoft.Data.SqlClient

Public Class PackingDBHelper

    Private Shared ReadOnly conString As String = ConfigurationManager.ConnectionStrings("Db_Server").ConnectionString


#Region "Project List"
        Public Shared Function GetProjectList() As List(Of PackingProjModel)
        Dim Projectlist As New List(Of PackingProjModel)
        Using con As New SqlConnection(conString)
            con.Open()
            Using cmd As New SqlCommand("SELECT * FROM V_PROJ_PACKING ORDER BY Proj_no", con)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Projectlist.Add(New PackingProjModel With {
                            .Proj_no = rdr("Proj_No").ToString(),
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


        Public Shared Function GetPackingJobs(proNo As String) As List(Of PackingModel)
            Dim PackingJobs As New List(Of PackingModel)

            Using con As New SqlConnection(conString)
            ' Add WHERE clause to use the parameter
            Dim query As String = "SELECT * FROM V_JOB_PACKING WHERE [Proj_no] = @BOMNo"
            Dim cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@BOMNo", proNo)

                Try
                    con.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                    Dim job As New PackingModel With {
                            .WID = reader("WID").ToString(),
                            .BOMNo = reader("Proj_no").ToString(),
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
                            .ReadyQty = If(IsDBNull(reader("Ready_Qty")), 0, Convert.ToInt32(reader("Ready_Qty"))),
                            .BalanceQty = If(IsDBNull(reader("Balance Qty")), 0, Convert.ToInt32(reader("Balance Qty"))),
                            .Status = reader("Status").ToString()
                        }
                    PackingJobs.Add(job)
                    End While
                Catch ex As Exception
                    MessageBox.Show("Error loading Packing jobs: " & ex.Message)
                End Try
            End Using

            Return PackingJobs
        End Function

#Region "Bundle List"


    Public Shared Function GetBundleList(proNo As String) As List(Of PackingHeaderModel)
        Dim BundleList As New List(Of PackingHeaderModel)

        Using con As New SqlConnection(conString)
            ' Add WHERE clause to use the parameter
            Dim query As String = "SELECT * FROM PACK_HEADER WHERE [Proj_no] = @ProjNo ORDER BY Proj_no"
            Dim cmd As New SqlCommand(query, con)
            cmd.Parameters.AddWithValue("@ProjNo", proNo)

            Try
                con.Open()
                Dim rdr As SqlDataReader = cmd.ExecuteReader()
                While rdr.Read()
                    Dim job As New PackingHeaderModel With {
                        .PackID = rdr("PackID").ToString(),
                        .PackEntryDate = If(IsDBNull(rdr("Pack_entry_date")), CType(Nothing, Date?), Convert.ToDateTime(rdr("Pack_entry_date"))),
                        .PackType = rdr("Pack_type").ToString(),
                        .ProjNo = rdr("Proj_no").ToString(),
                        .CreatedBy = rdr("Createdby").ToString()
                        }
                    BundleList.Add(job)
                End While
            Catch ex As Exception
                MessageBox.Show("Error loading Packing jobs: " & ex.Message)
            End Try
        End Using

        Return BundleList
    End Function

#End Region



#Region "Packtype"
    Public Shared Function GetPacktypeList() As List(Of PackingTypeModel)
        Dim PackingTypelist As New List(Of PackingTypeModel)
        Using con As New SqlConnection(conString)
            con.Open()
            Using cmd As New SqlCommand("SELECT * FROM Packtype", con)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        PackingTypelist.Add(New PackingTypeModel With {
                            .Packtype = rdr("Packtype").ToString()
                        })
                    End While
                End Using
            End Using
        End Using
        Return PackingTypelist
    End Function

#End Region



#Region "Packdetails"
    Public Shared Function GetPackEntryList(PackID As String) As List(Of PackEntryModel)
        Dim PackEntries As New List(Of PackEntryModel)

        Using con As New SqlConnection(conString)
            Dim query As String = "SELECT * FROM V_JOB_PACKeditems WHERE [PID] = @PID ORDER BY PID"
            Dim cmd As New SqlCommand(query, con)
            cmd.Parameters.AddWithValue("@PID", PackID)

            Try
                con.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read()
                    Dim entry As New PackEntryModel With {
                    .ID = Convert.ToInt32(reader("ID")),
                    .PID = reader("PID").ToString(),
                    .WID = reader("WID").ToString(),
                    .ProjNo = reader("Proj_no").ToString(),
                    .JCno = reader("JC_no").ToString(),
                    .enterDate = If(IsDBNull(reader("date")), Nothing, Convert.ToDateTime(reader("date"))),
                    .Description = reader("Description").ToString(),
                    .Qty = If(IsDBNull(reader("qty")), 0, Convert.ToInt32(reader("qty"))),
                    .UserID = reader("USERID").ToString()
                }
                    PackEntries.Add(entry)
                End While
            Catch ex As Exception
                MessageBox.Show("Error loading Pack Entries: " & ex.Message)
            End Try
        End Using

        Return PackEntries
    End Function

#End Region

    Public Shared Sub CreateBundle(Packtype As String, Projno As String, Createdby As String)
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand("sp_CreateBundle", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@Pack_type", Packtype)
                cmd.Parameters.AddWithValue("@Proj_no", Projno)
                cmd.Parameters.AddWithValue("@Createdby", Createdby)
                con.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub




    Public Shared Sub Additemstobundle(WID As String, Qty As Int32, PID As String, projno As String, JCno As String, createdBy As String)
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand("sp_Addtobundle", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@WID", WID)
                cmd.Parameters.AddWithValue("@Pack_type ", "")
                cmd.Parameters.AddWithValue("@Pack_entry_qty", Qty)
                cmd.Parameters.AddWithValue("@PackID ", PID)
                cmd.Parameters.AddWithValue("@Proj_no", Qty)
                cmd.Parameters.AddWithValue("@JC_no", JCno)
                cmd.Parameters.AddWithValue("@Created_by ", createdBy)
                con.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub




    Public Shared Sub Removeitems(ID As Int32)
        Using con As New SqlConnection(conString)
            Using cmd As New SqlCommand("sp_DeletePackEntry", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@PackEntryID", ID)
                con.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub







End Class
