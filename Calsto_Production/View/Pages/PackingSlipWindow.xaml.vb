Imports System.Windows
Imports System.Windows.Documents
Imports System.Windows.Media
Imports System.Printing
Imports System.Windows.Media.Imaging
Imports ZXing
Imports ZXing.Common
Imports ZXing.Windows.Compatibility
Imports ZXing.Rendering
Imports System.Drawing
Imports System.IO


Public Class PackingSlipWindow

    Private ReadOnly _bundle As PackingslipHeaderModel
    Private _flowDoc As FlowDocument

    Public Sub New(bundleData As PackingslipHeaderModel)
        InitializeComponent()
        _bundle = bundleData
        BuildPackingSlip()
    End Sub

    Private Sub BuildPackingSlip()
        ' --- FlowDocument setup ---
        Dim doc As New FlowDocument()
        doc.PageWidth = 794
        doc.PageHeight = 1123
        doc.PagePadding = New Thickness(50)
        doc.ColumnWidth = doc.PageWidth - doc.PagePadding.Left - doc.PagePadding.Right
        doc.FontFamily = New System.Windows.Media.FontFamily("Segoe UI")
        doc.FontSize = 11

        ' --- Pagination parameters ---
        Const rowsPerPage As Integer = 35
        Dim index As Integer = 1
        Dim totalRows = _bundle.Items.Count

        ' --- Add items table in chunks ---
        While index <= totalRows

            ' Pack ID header at top of each page
            doc.Blocks.Add(CreateTopHeader(_bundle))
            doc.Blocks.Add(CreatePackIDHeader(_bundle.PackID))



            ' Items table for this page
            Dim table As Table = CreateItemsTable(_bundle.Items.Skip(index - 1).Take(rowsPerPage).ToList(), index)
            doc.Blocks.Add(table)

            index += rowsPerPage
        End While

        ' --- Footer ---
        Dim footer As New Paragraph(New Run("Signature: ___________________________"))
        footer.Margin = New Thickness(0, 40, 0, 0)
        doc.Blocks.Add(footer)

        ' --- Store for preview ---
        _flowDoc = doc
        DocViewer.Document = _flowDoc
    End Sub
    Private Function CreateTopHeader(bundle As PackingslipHeaderModel) As Section
        Dim section As New Section()

        ' --- Header Table ---
        Dim headerTable As New Table()
        headerTable.Columns.Add(New TableColumn() With {.Width = New GridLength(150)}) ' Logo
        headerTable.Columns.Add(New TableColumn() With {.Width = New GridLength(350)}) ' Company + Title
        headerTable.Columns.Add(New TableColumn() With {.Width = New GridLength(150)}) ' QR Code

        Dim topHeaderGroup As New TableRowGroup()
        Dim row As New TableRow()

        ' Logo
        Dim logoCell As New TableCell()
        Dim logoImg As New System.Windows.Controls.Image() With {
        .Source = New BitmapImage(New Uri("pack://application:,,,/Resources/CA_Storage_Logo-removebg-preview.png")),
        .Width = 120,
        .Height = 60,
        .HorizontalAlignment = HorizontalAlignment.Left
    }
        logoCell.Blocks.Add(New BlockUIContainer(logoImg))
        row.Cells.Add(logoCell)

        ' Company Name
        Dim companyCell As New TableCell(New Paragraph(New Run("CRAFTSMAN STORAGE SOLUTIONS")) With {
        .FontWeight = FontWeights.ExtraBold,
        .FontSize = 22,
        .TextAlignment = TextAlignment.Center
    })
        row.Cells.Add(companyCell)

        ' QR Code (top-right corner)
        Dim qrCell As New TableCell()
        Dim qrImg As New System.Windows.Controls.Image() With {
        .Source = GenerateQRCodeImage(bundle.PackID, 100, 100),
        .Width = 100,
        .Height = 100,
        .HorizontalAlignment = HorizontalAlignment.Right
    }
        qrCell.Blocks.Add(New BlockUIContainer(qrImg))
        row.Cells.Add(qrCell)

        topHeaderGroup.Rows.Add(row)
        headerTable.RowGroups.Add(topHeaderGroup)
        section.Blocks.Add(headerTable)

        ' --- Customer Info Table ---
        Dim infoTable As New Table()
        infoTable.Columns.Add(New TableColumn() With {.Width = New GridLength(100)})
        infoTable.Columns.Add(New TableColumn() With {.Width = New GridLength(320)})
        infoTable.Columns.Add(New TableColumn() With {.Width = New GridLength(100)})
        infoTable.Columns.Add(New TableColumn() With {.Width = New GridLength(200)})

        Dim infoGroup As New TableRowGroup()

        ' Row 1
        Dim infoRow As New TableRow()
        infoRow.Cells.Add(New TableCell(New Paragraph(New Run("Customer:"))) With {.FontSize = 14, .TextAlignment = TextAlignment.Left})
        infoRow.Cells.Add(New TableCell(New Paragraph(New Run(bundle.Customer)) With {.FontSize = 16, .FontWeight = FontWeights.Bold}))
        infoRow.Cells.Add(New TableCell(New Paragraph(New Run("Project no.:"))) With {.FontSize = 14, .TextAlignment = TextAlignment.Left})
        infoRow.Cells.Add(New TableCell(New Paragraph(New Run(bundle.ProjectNo)) With {.FontSize = 16, .FontWeight = FontWeights.Bold}))
        infoGroup.Rows.Add(infoRow)

        ' Row 2
        Dim infoRow2 As New TableRow()
        infoRow2.Cells.Add(New TableCell(New Paragraph(New Run("Location:"))) With {.FontSize = 14, .TextAlignment = TextAlignment.Left})
        infoRow2.Cells.Add(New TableCell(New Paragraph(New Run(bundle.Location)) With {.FontSize = 16, .FontWeight = FontWeights.Bold}))
        infoRow2.Cells.Add(New TableCell(New Paragraph(New Run("Package type:"))) With {.FontSize = 16, .TextAlignment = TextAlignment.Left})
        infoRow2.Cells.Add(New TableCell(New Paragraph(New Run(bundle.PackType)) With {.FontSize = 16, .FontWeight = FontWeights.Bold}))
        infoGroup.Rows.Add(infoRow2)

        ' Row 3
        Dim infoRow3 As New TableRow()
        infoRow3.Cells.Add(New TableCell(New Paragraph(New Run("System:"))) With {.FontSize = 14, .TextAlignment = TextAlignment.Left})
        infoRow3.Cells.Add(New TableCell(New Paragraph(New Run(bundle.SystemType)) With {.FontSize = 14, .FontWeight = FontWeights.Bold}))
        infoRow3.Cells.Add(New TableCell(New Paragraph(New Run(""))))
        infoRow3.Cells.Add(New TableCell(New Paragraph(New Run(""))))
        infoGroup.Rows.Add(infoRow3)

        infoTable.RowGroups.Add(infoGroup)
        section.Blocks.Add(infoTable)

        Return section
    End Function



    ' --- Pack ID header for every page ---
    Private Function CreatePackIDHeader(packID As String) As Section
        Dim section As New Section()
        ' Pack ID text
        Dim para As New Paragraph(New Run("PID: " & packID)) With {
        .FontSize = 20,
        .FontWeight = FontWeights.Bold,
        .TextAlignment = TextAlignment.Center
    }
        section.Blocks.Add(para)

        Return section
    End Function



    Private Function GenerateQRCodeImage(data As String, width As Integer, height As Integer) As BitmapImage
        ' ZXing writer
        Dim writer As New BarcodeWriter() With {
    .Format = BarcodeFormat.QR_CODE,
    .Options = New EncodingOptions With {
        .Width = width,
        .Height = height,
        .Margin = 0
    },
    .Renderer = New BitmapRenderer()
} ' <- important


        Dim bitmap As Bitmap = writer.Write(data)

        ' Convert Bitmap to BitmapImage (WPF)
        Using ms As New MemoryStream()
            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
            ms.Position = 0
            Dim bmpImage As New BitmapImage()
            bmpImage.BeginInit()
            bmpImage.StreamSource = ms
            bmpImage.CacheOption = BitmapCacheOption.OnLoad
            bmpImage.EndInit()
            Return bmpImage
        End Using
    End Function




    ' --- Items table ---
    Private Function CreateItemsTable(items As List(Of PackingItemModel), startIndex As Integer) As Table
        Dim table As New Table() With {
            .CellSpacing = 0,
            .BorderBrush = System.Windows.Media.Brushes.Black,
            .BorderThickness = New Thickness(1)
        }

        ' Columns
        table.Columns.Add(New TableColumn() With {.Width = New GridLength(50)})
        table.Columns.Add(New TableColumn() With {.Width = New GridLength(200)})
        table.Columns.Add(New TableColumn() With {.Width = New GridLength(330)})
        table.Columns.Add(New TableColumn() With {.Width = New GridLength(100)})

        ' Header row
        Dim headerRow As New TableRow()
        headerRow.Cells.Add(MakeHeaderCell("S.No"))
        headerRow.Cells.Add(MakeHeaderCell("Part Number"))
        headerRow.Cells.Add(MakeHeaderCell("Description"))
        headerRow.Cells.Add(MakeHeaderCell("Qty"))

        Dim headerGroup As New TableRowGroup()
        headerGroup.Rows.Add(headerRow)
        table.RowGroups.Add(headerGroup)

        ' Body rows
        Dim bodyGroup As New TableRowGroup()
        Dim index As Integer = startIndex
        For Each itm In items
            Dim row As New TableRow()
            ' Alternating row color
            If index Mod 2 = 0 Then row.Background = System.Windows.Media.Brushes.LightGray

            row.Cells.Add(MakeBodyCell(index.ToString()))
            row.Cells.Add(MakeBodyCell(itm.PartNumber))
            row.Cells.Add(MakeBodyCell(itm.Description))
            row.Cells.Add(MakeBodyCell(itm.Qty.ToString()))

            bodyGroup.Rows.Add(row)
            index += 1
        Next
        table.RowGroups.Add(bodyGroup)

        Return table
    End Function

    ' --- Header Cell ---
    Private Function MakeHeaderCell(text As String) As TableCell
        Return New TableCell(New Paragraph(New Run(text))) With {
            .FontWeight = FontWeights.Bold,
            .TextAlignment = TextAlignment.Center,
            .Padding = New Thickness(3),
            .BorderBrush = System.Windows.Media.Brushes.Black,
            .BorderThickness = New Thickness(1)
        }
    End Function

    ' --- Body Cell ---
    Private Function MakeBodyCell(text As String) As TableCell
        Return New TableCell(New Paragraph(New Run(text))) With {
            .TextAlignment = TextAlignment.Left,
            .Padding = New Thickness(2),
            .BorderBrush = System.Windows.Media.Brushes.Black,
            .BorderThickness = New Thickness(1)
        }
    End Function

End Class
