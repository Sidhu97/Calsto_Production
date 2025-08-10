Imports System
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media.Imaging
Imports QRCoder
Imports System.Drawing
Imports System.IO

Namespace Calsto_Production
    Partial Public Class PackingSlip
        Inherits Window

        ' Properties for binding
        Public Property CustomerName As String
        Public Property CustomerLocation As String
        Public Property ProjectNo As String
        Public Property PreparedDate As DateTime
        Public Property QualityDate As DateTime
        Public Property PackageID As String

        Public Sub New()
            InitializeComponent()
            Me.DataContext = Me
        End Sub

        Public Sub SetQRCode(qrData As String)
            Dim qrGenerator As New QRCodeGenerator()
            Dim qrCodeData = qrGenerator.CreateQrCode(qrData, QRCodeGenerator.ECCLevel.Q)
            Dim qrCode As New QRCode(qrCodeData)
            Dim qrBitmap As Bitmap = qrCode.GetGraphic(20)

            ' Convert Bitmap to BitmapImage
            Using memory As New MemoryStream()
                qrBitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png)
                memory.Position = 0
                Dim bitmapImage As New BitmapImage()
                bitmapImage.BeginInit()
                bitmapImage.StreamSource = memory
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad
                bitmapImage.EndInit()
                imgQRCode.Source = bitmapImage
            End Using
        End Sub

        Private Sub PrintButton_Click(sender As Object, e As RoutedEventArgs)
            Dim printDialog As New PrintDialog()
            If printDialog.ShowDialog() = True Then
                printDialog.PrintVisual(Me, "Packing Slip")
            End If
        End Sub
    End Class
End Namespace