Imports System.Windows
Imports System.Windows.Documents
Imports System.Windows.Media
Imports System.Printing
Imports System.Windows.Media.Imaging

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
        doc.FontFamily = New FontFamily("Segoe UI")
        doc.FontSize = 12

        ' --- Add top logo and company info (only first page) ---
        ' --- Pagination parameters ---
        Const rowsPerPage As Integer = 30
        Dim index As Integer = 1
        Dim totalRows = _bundle.Items.Count

        ' --- Add items table in chunks ---
        While index <= totalRows
            ' Pack ID header at top of each page
            doc.Blocks.Add(CreateTopHeader())
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
    Private Function CreateTopHeader() As Section
        Dim section As New Section()

        ' Header Table
        Dim headerTable As New Table()
        headerTable.Columns.Add(New TableColumn() With {.Width = New GridLength(200)})
        headerTable.Columns.Add(New TableColumn() With {.Width = New GridLength(300)})
        headerTable.Columns.Add(New TableColumn() With {.Width = New GridLength(200)})

        Dim topheaderGroup As New TableRowGroup()
        Dim row As New TableRow()

        ' Logo
        Dim logoCell As New TableCell()
        Dim logoImg As New Image() With {
            .Source = New BitmapImage(New Uri("pack://application:,,,/Resources/CA_Storage_Logo-removebg-preview.png")),
            .Width = 100,
            .Height = 50
        }
        logoCell.Blocks.Add(New BlockUIContainer(logoImg))
        row.Cells.Add(logoCell)

        ' Company Name
        Dim companyCell As New TableCell(New Paragraph(New Run("CRAFTSMAN STORAGE SOLUTIONS")) With {
            .FontWeight = FontWeights.Bold,
            .FontSize = 16,
            .TextAlignment = TextAlignment.Center
        })
        row.Cells.Add(companyCell)

        ' Barcode placeholder
        Dim barcodeCell As New TableCell()
        Dim barcodeImg As New Image() With {
            .Source = New BitmapImage(New Uri("pack://application:,,,/Resources/CA_Storage_Logo-removebg-preview.png")),
            .Width = 150,
            .Height = 50
        }
        barcodeCell.Blocks.Add(New BlockUIContainer(barcodeImg))
        row.Cells.Add(barcodeCell)

        topheaderGroup.Rows.Add(row)
        headerTable.RowGroups.Add(topheaderGroup)
        section.Blocks.Add(headerTable)

        ' Customer info table
        Dim infoTable As New Table()
        infoTable.Columns.Add(New TableColumn() With {.Width = New GridLength(100)})
        infoTable.Columns.Add(New TableColumn() With {.Width = New GridLength(300)})
        infoTable.Columns.Add(New TableColumn() With {.Width = New GridLength(100)})
        infoTable.Columns.Add(New TableColumn() With {.Width = New GridLength(200)})

        Dim infoGroup As New TableRowGroup()

        ' Row 1
        Dim infoRow As New TableRow()
        infoRow.Cells.Add(New TableCell(New Paragraph(New Run("Customer"))) With {.TextAlignment = TextAlignment.Left})
        infoRow.Cells.Add(New TableCell(New Paragraph(New Bold(New Run("Flipkart, Uluberia"))) With {.FontSize = 15, .TextAlignment = TextAlignment.Left}))
        infoRow.Cells.Add(New TableCell(New Paragraph(New Bold(New Run("Project no."))) With {.TextAlignment = TextAlignment.Right}))
        infoRow.Cells.Add(New TableCell(New Paragraph(New Run("CSS-3001"))))
        infoGroup.Rows.Add(infoRow)

        ' Row 2
        Dim infoRow2 As New TableRow()
        infoRow2.Cells.Add(New TableCell(New Paragraph(New Bold(New Run("Location"))) With {.TextAlignment = TextAlignment.Left}))
        infoRow2.Cells.Add(New TableCell(New Paragraph(New Run("Kolkata"))))
        infoRow2.Cells.Add(New TableCell(New Paragraph(New Bold(New Run("Package type"))) With {.TextAlignment = TextAlignment.Right}))
        infoRow2.Cells.Add(New TableCell(New Paragraph(New Run("Wooden Pallet"))))
        infoGroup.Rows.Add(infoRow2)

        ' Row 3
        Dim infoRow3 As New TableRow()
        infoRow3.Cells.Add(New TableCell(New Paragraph(New Bold(New Run("System"))) With {.TextAlignment = TextAlignment.Left}))
        infoRow3.Cells.Add(New TableCell(New Paragraph(New Run("SPR/SHELVING"))))
        infoRow3.Cells.Add(New TableCell(New Paragraph(New Run(""))))
        infoRow3.Cells.Add(New TableCell(New Paragraph(New Run(""))))
        infoGroup.Rows.Add(infoRow3)

        infoTable.RowGroups.Add(infoGroup)
        section.Blocks.Add(infoTable)

        Return section

    End Function

    ' --- Pack ID header for every page ---
    Private Function CreatePackIDHeader(packID As String) As Paragraph
        Dim header As New Paragraph(New Run("Pack ID: " & packID)) With {
            .FontSize = 14,
            .FontWeight = FontWeights.Bold,
            .TextAlignment = TextAlignment.Center,
            .Margin = New Thickness(0, 0, 0, 10)
        }
        Return header
    End Function

    ' --- Items table ---
    Private Function CreateItemsTable(items As List(Of PackingItemModel), startIndex As Integer) As Table
        Dim table As New Table() With {
            .CellSpacing = 0,
            .BorderBrush = Brushes.Black,
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
            If index Mod 2 = 0 Then row.Background = Brushes.LightGray

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
            .Padding = New Thickness(4),
            .BorderBrush = Brushes.Black,
            .BorderThickness = New Thickness(1)
        }
    End Function

    ' --- Body Cell ---
    Private Function MakeBodyCell(text As String) As TableCell
        Return New TableCell(New Paragraph(New Run(text))) With {
            .TextAlignment = TextAlignment.Left,
            .Padding = New Thickness(4),
            .BorderBrush = Brushes.Black,
            .BorderThickness = New Thickness(1)
        }
    End Function

End Class
