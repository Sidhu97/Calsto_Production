Imports System.Configuration
Imports System.Data
Imports Microsoft.Data.SqlClient

Public Class PackingSlipDBHelper

    Private Shared ReadOnly conString As String = ConfigurationManager.ConnectionStrings("Db_Server").ConnectionString

#Region "Packing Details"

    ' Return a single header because PackID is unique
    Public Shared Function GetPackingHeader(PackID As String) As PackingslipHeaderModel
        Dim header As PackingslipHeaderModel = Nothing

        Using con As New SqlConnection(conString)
            Dim query As String = "SELECT * FROM V_Slip_BundleList WHERE PackID = @PackID"
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@PackID", PackID)

                Try
                    con.Open()
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            header = New PackingslipHeaderModel With {
                                .PackID = reader("PackID").ToString(),
                                .ProjectNo = reader("ProjectNo").ToString(),
                                .Customer = reader("Customer").ToString(),
                                .Location = reader("Location").ToString(),
                                .SystemType = reader("SystemType").ToString(),
                                .PackType = reader("PackType").ToString(),
                                .PackedBy = reader("PackedBy").ToString(),
                                .PackDate = If(IsDBNull(reader("PackDate")), DateTime.MinValue, Convert.ToDateTime(reader("PackDate")))
                            }
                        End If
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error loading packing header: " & ex.Message)
                End Try
            End Using
        End Using

        Return header
    End Function

#End Region

#Region "Bundle Item List"

    Public Shared Function GetPackingItems(PackID As String) As List(Of PackingItemModel)
        Dim items As New List(Of PackingItemModel)

        Using con As New SqlConnection(conString)
            Dim query As String = "
            SELECT WID,
                   [Part Number] AS PartNumber,
                   Description,
                   QTY
            FROM V_Slip_ItemList
            WHERE PackID = @PackID"

            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@PackID", PackID)

                Try
                    con.Open()
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim item As New PackingItemModel With {
                                .WID = reader("WID").ToString(),
                                .PartNumber = reader("PartNumber").ToString(),
                                .Description = reader("Description").ToString(),
                                .Qty = Convert.ToInt32(reader("QTY"))
                            }
                            items.Add(item)
                        End While
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error loading packing items: " & ex.Message)
                End Try
            End Using
        End Using

        Return items
    End Function

#End Region

End Class
