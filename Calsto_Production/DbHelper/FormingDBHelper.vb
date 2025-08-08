Imports System.Configuration
Imports System.Data
Imports Microsoft.Data.SqlClient

Public Class FormingDBHelper

    Private Shared ReadOnly conString As String = ConfigurationManager.ConnectionStrings("Db_Server").ConnectionString


#Region "Get Forming Jobs"
    Public Shared Function GetFormingJobs() As List(Of FormingModel)
        Dim formingJobs As New List(Of FormingModel)

        Using con As New SqlConnection(conString)
            Dim query As String = "SELECT * FROM V_JOB_FORMING"
            Dim cmd As New SqlCommand(query, con)

            Try
                con.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read()
                    Dim job As New FormingModel With {
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
                        .DispatchDate = If(IsDBNull(reader("Dispatch Date")), CType(Nothing, Date?), Convert.ToDateTime(reader("Dispatch Date"))),
                        .DueDate = If(IsDBNull(reader("Due Date")), 0, Convert.ToInt32(reader("Due Date")))
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

End Class
