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
            Dim query As String = "SELECT * FROM PACK_ENTRY WHERE [PackID] = @PackID ORDER BY PackID"
            Dim cmd As New SqlCommand(query, con)
            cmd.Parameters.AddWithValue("@PackID", PackID)

            Try
                con.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read()
                    Dim entry As New PackEntryModel With {
                    .WID = reader("WID").ToString(),
                    .PackEntryDate = If(IsDBNull(reader("Pack_entry_date")), Nothing, Convert.ToDateTime(reader("Pack_entry_date"))),
                    .PackType = reader("Pack_type").ToString(),
                    .PackEntryQty = If(IsDBNull(reader("Pack_entry_qty")), 0, Convert.ToInt32(reader("Pack_entry_qty"))),
                    .PackID = reader("PackID").ToString(),
                    .PackEntryID = Convert.ToInt32(reader("PackEntryID")),
                    .ProjNo = reader("Proj_no").ToString(),
                    .DispatchLocation = reader("Dispatch_Location").ToString(),
                    .VehNo = reader("veh_no").ToString(),
                    .DriverNo = reader("Driver_No").ToString(),
                    .DispatchedDate = If(IsDBNull(reader("Dispatched_date")), Nothing, Convert.ToDateTime(reader("Dispatched_date"))),
                    .Status = reader("Status").ToString(),
                    .RecivedDate = If(IsDBNull(reader("Recived_date")), Nothing, Convert.ToDateTime(reader("Recived_date"))),
                    .UserID = reader("USERID").ToString(),
                    .SystemName = reader("SYSTEMNAME").ToString(),
                    .UnLoadUser = reader("UnLoadUser").ToString(),
                    .AttachFile = reader("AttachFile").ToString(),
                    .VehicleSerialNo = reader("vehicle_SerialNo").ToString(),
                    .Reason = reader("Reason").ToString()
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

End Class
