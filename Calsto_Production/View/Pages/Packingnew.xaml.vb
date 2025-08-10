Imports System
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports QRCoder
Imports System.Drawing
Imports System.Windows.Media.Imaging

Namespace Calsto_Production
    Partial Public Class Packingnew
        Inherits Page

        ' Collections
        Private projects As New ObservableCollection(Of Project)()
        Private availableComponents As New ObservableCollection(Of Component)()
        Private bundles As New ObservableCollection(Of Bundle)()

        Private selectedProject As Project
        Private selectedBundle As Bundle

        Public Sub New()
            InitializeComponent()

            ' Load only projects on start; keep components and bundles empty
            InitializeProjects()

            ' Bind UI elements
            cmbProject.ItemsSource = projects
            cmbProject.DisplayMemberPath = "Name"
            cmbBundleType.ItemsSource = New List(Of String) From {"Bundle", "Box", "Pallet"}

            dgAvailableComponents.ItemsSource = availableComponents
            dgBundleList.ItemsSource = bundles
            dgBundleComponents.ItemsSource = Nothing

            ' Set initial UI state after all bindings are done
            ' This ensures everything starts in the correct disabled state
            UpdateUIState()
        End Sub

        Private Sub InitializeProjects()
            projects.Add(New Project With {.Name = "CSS-3001"})
            projects.Add(New Project With {.Name = "CSS-3002"})
            projects.Add(New Project With {.Name = "CSS-3003"})
        End Sub

        Private Sub UpdateUIState()
            ' State flags
            Dim hasProjectSelected As Boolean = selectedProject IsNot Nothing
            Dim hasBundleSelected As Boolean = selectedBundle IsNot Nothing
            Dim hasBundleItems As Boolean = hasBundleSelected AndAlso
                                     selectedBundle.Components IsNot Nothing AndAlso
                                     selectedBundle.Components.Count > 0
            Dim hasBundleTypeSelected As Boolean = cmbBundleType.SelectedItem IsNot Nothing


            cmbProject.IsEnabled = True


            cmbBundleType.IsEnabled = hasProjectSelected


            btnCreateBundle.IsEnabled = hasProjectSelected AndAlso hasBundleTypeSelected


            btnSave.IsEnabled = bundles.Count > 0


            dgAvailableComponents.IsEnabled = hasBundleSelected


            btnAddToBundle.IsEnabled = hasBundleSelected


            btnRemoveFromBundle.IsEnabled = hasBundleSelected AndAlso hasBundleItems


            btnPrintPreview.IsEnabled = hasBundleItems

            If bundles.Count = 0 Then
                dgBundleList.UnselectAll()
                selectedBundle = Nothing
            End If
            If Not dgAvailableComponents.IsEnabled Then
                dgAvailableComponents.UnselectAll()
            End If
            If Not hasBundleSelected Then
                dgBundleComponents.UnselectAll()
            End If
        End Sub
        Private Sub cmbProject_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            selectedProject = TryCast(cmbProject.SelectedItem, Project)
            If selectedProject Is Nothing Then Return
            availableComponents.Clear()
            bundles.Clear()
            LoadSampleComponentsForProject(selectedProject.Name)
            UpdateUIState()
        End Sub

        Private Sub cmbBundleType_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            UpdateUIState()
        End Sub

        Private Sub LoadSampleComponentsForProject(projectName As String)
            If projectName = "CSS-3001" Then
                availableComponents.Add(New Component With {
                    .SrNo = 1, .WID = "WID101", .CustomerName = "Cust11", .CustomerLocation = "Loc11",
                    .ProjectNo = projectName, .Description = "Socket Head Cap Screw", .QtyAvailable = 25,
                    .Color = "Black", .ReadyToPackQty = 20, .PendingQty = 5, .BalanceQty = 0,
                    .Location = "A1-1", .Status = "Available", .PartNumber = "PN101", .Remarks = "Checked and approved"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 2, .WID = "WID102", .CustomerName = "Cust12", .CustomerLocation = "Loc12",
                    .ProjectNo = projectName, .Description = "Flat Washer", .QtyAvailable = 15,
                    .Color = "Silver", .ReadyToPackQty = 15, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "A1-2", .Status = "Available", .PartNumber = "PN102", .Remarks = "Ready for dispatch"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 3, .WID = "WID103", .CustomerName = "Cust13", .CustomerLocation = "Loc13",
                    .ProjectNo = projectName, .Description = "Hex Nut", .QtyAvailable = 50,
                    .Color = "Zinc", .ReadyToPackQty = 45, .PendingQty = 5, .BalanceQty = 0,
                    .Location = "A1-3", .Status = "Available", .PartNumber = "PN103", .Remarks = "Pending QA check"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 4, .WID = "WID104", .CustomerName = "Cust14", .CustomerLocation = "Loc14",
                    .ProjectNo = projectName, .Description = "Spring Pin", .QtyAvailable = 80,
                    .Color = "Steel", .ReadyToPackQty = 80, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "A1-4", .Status = "Available", .PartNumber = "PN104", .Remarks = "Good condition"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 5, .WID = "WID105", .CustomerName = "Cust15", .CustomerLocation = "Loc15",
                    .ProjectNo = projectName, .Description = "Lock Nut", .QtyAvailable = 12,
                    .Color = "Yellow", .ReadyToPackQty = 10, .PendingQty = 2, .BalanceQty = 0,
                    .Location = "A1-5", .Status = "Available", .PartNumber = "PN105", .Remarks = "Pending material"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 6, .WID = "WID106", .CustomerName = "Cust16", .CustomerLocation = "Loc16",
                    .ProjectNo = projectName, .Description = "Washer", .QtyAvailable = 15,
                    .Color = "Black", .ReadyToPackQty = 12, .PendingQty = 3, .BalanceQty = 0,
                    .Location = "A1-6", .Status = "Available", .PartNumber = "PN106", .Remarks = "Hold for inspection"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 7, .WID = "WID107", .CustomerName = "Cust17", .CustomerLocation = "Loc17",
                    .ProjectNo = projectName, .Description = "Threaded Rod", .QtyAvailable = 7,
                    .Color = "Gray", .ReadyToPackQty = 7, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "A1-7", .Status = "Available", .PartNumber = "PN107", .Remarks = "Ready to ship"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 8, .WID = "WID108", .CustomerName = "Cust18", .CustomerLocation = "Loc18",
                    .ProjectNo = projectName, .Description = "Wing Nut", .QtyAvailable = 9,
                    .Color = "Green", .ReadyToPackQty = 9, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "A1-8", .Status = "Available", .PartNumber = "PN108", .Remarks = "Packaged"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 9, .WID = "WID109", .CustomerName = "Cust19", .CustomerLocation = "Loc19",
                    .ProjectNo = projectName, .Description = "Cotter Pin", .QtyAvailable = 25,
                    .Color = "Bronze", .ReadyToPackQty = 20, .PendingQty = 5, .BalanceQty = 0,
                    .Location = "A1-9", .Status = "Available", .PartNumber = "PN109", .Remarks = "Final check needed"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 10, .WID = "WID110", .CustomerName = "Cust20", .CustomerLocation = "Loc20",
                    .ProjectNo = projectName, .Description = "Eye Bolt", .QtyAvailable = 6,
                    .Color = "Gold", .ReadyToPackQty = 6, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "A1-10", .Status = "Available", .PartNumber = "PN110", .Remarks = "Ready to go"
                })
                availableComponents.Add(New Component With {
            .SrNo = 6, .WID = "WID106", .CustomerName = "Cust16", .CustomerLocation = "Loc16",
            .ProjectNo = projectName, .Description = "Washer", .QtyAvailable = 15,
            .Color = "Black", .ReadyToPackQty = 12, .PendingQty = 3, .BalanceQty = 0,
            .Location = "A1-6", .Status = "Available", .PartNumber = "PN106", .Remarks = "Hold for inspection"
        })
                availableComponents.Add(New Component With {
                    .SrNo = 7, .WID = "WID107", .CustomerName = "Cust17", .CustomerLocation = "Loc17",
                    .ProjectNo = projectName, .Description = "Threaded Rod", .QtyAvailable = 7,
                    .Color = "Gray", .ReadyToPackQty = 7, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "A1-7", .Status = "Available", .PartNumber = "PN107", .Remarks = "Ready to ship"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 8, .WID = "WID108", .CustomerName = "Cust18", .CustomerLocation = "Loc18",
                    .ProjectNo = projectName, .Description = "Wing Nut", .QtyAvailable = 9,
                    .Color = "Green", .ReadyToPackQty = 9, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "A1-8", .Status = "Available", .PartNumber = "PN108", .Remarks = "Packaged"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 9, .WID = "WID109", .CustomerName = "Cust19", .CustomerLocation = "Loc19",
                    .ProjectNo = projectName, .Description = "Cotter Pin", .QtyAvailable = 25,
                    .Color = "Bronze", .ReadyToPackQty = 20, .PendingQty = 5, .BalanceQty = 0,
                    .Location = "A1-9", .Status = "Available", .PartNumber = "PN109", .Remarks = "Final check needed"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 10, .WID = "WID110", .CustomerName = "Cust20", .CustomerLocation = "Loc20",
                    .ProjectNo = projectName, .Description = "Eye Bolt", .QtyAvailable = 6,
                    .Color = "Gold", .ReadyToPackQty = 6, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "A1-10", .Status = "Available", .PartNumber = "PN110", .Remarks = "Ready to go"
                })
                availableComponents.Add(New Component With {
            .SrNo = 6, .WID = "WID106", .CustomerName = "Cust16", .CustomerLocation = "Loc16",
            .ProjectNo = projectName, .Description = "Washer", .QtyAvailable = 15,
            .Color = "Black", .ReadyToPackQty = 12, .PendingQty = 3, .BalanceQty = 0,
            .Location = "A1-6", .Status = "Available", .PartNumber = "PN106", .Remarks = "Hold for inspection"
        })
                availableComponents.Add(New Component With {
                    .SrNo = 7, .WID = "WID107", .CustomerName = "Cust17", .CustomerLocation = "Loc17",
                    .ProjectNo = projectName, .Description = "Threaded Rod", .QtyAvailable = 7,
                    .Color = "Gray", .ReadyToPackQty = 7, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "A1-7", .Status = "Available", .PartNumber = "PN107", .Remarks = "Ready to ship"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 8, .WID = "WID108", .CustomerName = "Cust18", .CustomerLocation = "Loc18",
                    .ProjectNo = projectName, .Description = "Wing Nut", .QtyAvailable = 9,
                    .Color = "Green", .ReadyToPackQty = 9, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "A1-8", .Status = "Available", .PartNumber = "PN108", .Remarks = "Packaged"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 9, .WID = "WID109", .CustomerName = "Cust19", .CustomerLocation = "Loc19",
                    .ProjectNo = projectName, .Description = "Cotter Pin", .QtyAvailable = 25,
                    .Color = "Bronze", .ReadyToPackQty = 20, .PendingQty = 5, .BalanceQty = 0,
                    .Location = "A1-9", .Status = "Available", .PartNumber = "PN109", .Remarks = "Final check needed"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 10, .WID = "WID110", .CustomerName = "Cust20", .CustomerLocation = "Loc20",
                    .ProjectNo = projectName, .Description = "Eye Bolt", .QtyAvailable = 6,
                    .Color = "Gold", .ReadyToPackQty = 6, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "A1-10", .Status = "Available", .PartNumber = "PN110", .Remarks = "Ready to go"
                })
            End If

            If projectName = "CSS-3002" Then
                availableComponents.Add(New Component With {
                    .SrNo = 1, .WID = "WID001", .CustomerName = "Cust1", .CustomerLocation = "Loc1",
                    .ProjectNo = projectName, .Description = "Hexagonal Bolt", .QtyAvailable = 10,
                    .Color = "Red", .ReadyToPackQty = 10, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "Shelf1", .Status = "Available", .PartNumber = "PN001", .Remarks = "Remark1"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 2, .WID = "WID002", .CustomerName = "Cust2", .CustomerLocation = "Loc2",
                    .ProjectNo = projectName, .Description = "Flat Washer", .QtyAvailable = 5,
                    .Color = "Blue", .ReadyToPackQty = 5, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "Shelf2", .Status = "Available", .PartNumber = "PN002", .Remarks = "Remark2"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 3, .WID = "WID003", .CustomerName = "Cust3", .CustomerLocation = "Loc3",
                    .ProjectNo = projectName, .Description = "Nylon Washer", .QtyAvailable = 20,
                    .Color = "White", .ReadyToPackQty = 15, .PendingQty = 5, .BalanceQty = 0,
                    .Location = "Shelf3", .Status = "Available", .PartNumber = "PN003", .Remarks = "Remark3"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 4, .WID = "WID004", .CustomerName = "Cust4", .CustomerLocation = "Loc4",
                    .ProjectNo = projectName, .Description = "Spring Pin", .QtyAvailable = 8,
                    .Color = "Silver", .ReadyToPackQty = 8, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "Shelf4", .Status = "Available", .PartNumber = "PN004", .Remarks = "Remark4"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 5, .WID = "WID005", .CustomerName = "Cust5", .CustomerLocation = "Loc5",
                    .ProjectNo = projectName, .Description = "Lock Nut", .QtyAvailable = 12,
                    .Color = "Yellow", .ReadyToPackQty = 10, .PendingQty = 2, .BalanceQty = 0,
                    .Location = "Shelf5", .Status = "Available", .PartNumber = "PN005", .Remarks = "Remark5"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 6, .WID = "WID006", .CustomerName = "Cust6", .CustomerLocation = "Loc6",
                    .ProjectNo = projectName, .Description = "Socket Head Cap Screw", .QtyAvailable = 15,
                    .Color = "Black", .ReadyToPackQty = 12, .PendingQty = 3, .BalanceQty = 0,
                    .Location = "Shelf6", .Status = "Available", .PartNumber = "PN006", .Remarks = "Remark6"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 7, .WID = "WID007", .CustomerName = "Cust7", .CustomerLocation = "Loc7",
                    .ProjectNo = projectName, .Description = "Threaded Rod", .QtyAvailable = 7,
                    .Color = "Gray", .ReadyToPackQty = 7, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "Shelf7", .Status = "Available", .PartNumber = "PN007", .Remarks = "Remark7"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 8, .WID = "WID008", .CustomerName = "Cust8", .CustomerLocation = "Loc8",
                    .ProjectNo = projectName, .Description = "Wing Nut", .QtyAvailable = 9,
                    .Color = "Green", .ReadyToPackQty = 9, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "Shelf8", .Status = "Available", .PartNumber = "PN008", .Remarks = "Remark8"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 9, .WID = "WID009", .CustomerName = "Cust9", .CustomerLocation = "Loc9",
                    .ProjectNo = projectName, .Description = "Cotter Pin", .QtyAvailable = 25,
                    .Color = "Bronze", .ReadyToPackQty = 20, .PendingQty = 5, .BalanceQty = 0,
                    .Location = "Shelf9", .Status = "Available", .PartNumber = "PN009", .Remarks = "Remark9"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 10, .WID = "WID010", .CustomerName = "Cust10", .CustomerLocation = "Loc10",
                    .ProjectNo = projectName, .Description = "Eye Bolt", .QtyAvailable = 6,
                    .Color = "Gold", .ReadyToPackQty = 6, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "Shelf10", .Status = "Available", .PartNumber = "PN010", .Remarks = "Remark10"
                })
            End If

            If projectName = "CSS-3003" Then
                availableComponents.Add(New Component With {
                    .SrNo = 1, .WID = "WID301", .CustomerName = "Cust21", .CustomerLocation = "Loc21",
                    .ProjectNo = projectName, .Description = "Stud Bolt", .QtyAvailable = 15,
                    .Color = "Stainless Steel", .ReadyToPackQty = 15, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "B1-1", .Status = "Available", .PartNumber = "PN301", .Remarks = "Initial stock"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 2, .WID = "WID302", .CustomerName = "Cust22", .CustomerLocation = "Loc22",
                    .ProjectNo = projectName, .Description = "Machine Screw", .QtyAvailable = 30,
                    .Color = "Silver", .ReadyToPackQty = 25, .PendingQty = 5, .BalanceQty = 0,
                    .Location = "B1-2", .Status = "Available", .PartNumber = "PN302", .Remarks = "Awaiting final inspection"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 3, .WID = "WID303", .CustomerName = "Cust23", .CustomerLocation = "Loc23",
                    .ProjectNo = projectName, .Description = "Threaded Insert", .QtyAvailable = 20,
                    .Color = "Brass", .ReadyToPackQty = 20, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "B1-3", .Status = "Available", .PartNumber = "PN303", .Remarks = "Complete"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 4, .WID = "WID304", .CustomerName = "Cust24", .CustomerLocation = "Loc24",
                    .ProjectNo = projectName, .Description = "U-Bolt", .QtyAvailable = 10,
                    .Color = "Zinc Plated", .ReadyToPackQty = 10, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "B1-4", .Status = "Available", .PartNumber = "PN304", .Remarks = "Ready to pack"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 5, .WID = "WID305", .CustomerName = "Cust25", .CustomerLocation = "Loc25",
                    .ProjectNo = projectName, .Description = "Circlip", .QtyAvailable = 45,
                    .Color = "Black Oxide", .ReadyToPackQty = 40, .PendingQty = 5, .BalanceQty = 0,
                    .Location = "B1-5", .Status = "Available", .PartNumber = "PN305", .Remarks = "Waiting for customer approval"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 6, .WID = "WID306", .CustomerName = "Cust26", .CustomerLocation = "Loc26",
                    .ProjectNo = projectName, .Description = "Woodruff Key", .QtyAvailable = 18,
                    .Color = "Steel", .ReadyToPackQty = 18, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "B1-6", .Status = "Available", .PartNumber = "PN306", .Remarks = "QC passed"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 7, .WID = "WID307", .CustomerName = "Cust27", .CustomerLocation = "Loc27",
                    .ProjectNo = projectName, .Description = "Spring Washer", .QtyAvailable = 25,
                    .Color = "Gray", .ReadyToPackQty = 25, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "B1-7", .Status = "Available", .PartNumber = "PN307", .Remarks = "Ready to ship"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 8, .WID = "WID308", .CustomerName = "Cust28", .CustomerLocation = "Loc28",
                    .ProjectNo = projectName, .Description = "Rivnut", .QtyAvailable = 12,
                    .Color = "Aluminum", .ReadyToPackQty = 12, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "B1-8", .Status = "Available", .PartNumber = "PN308", .Remarks = "Ready for shipment"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 9, .WID = "WID309", .CustomerName = "Cust29", .CustomerLocation = "Loc29",
                    .ProjectNo = projectName, .Description = "Clevis Pin", .QtyAvailable = 7,
                    .Color = "Bronze", .ReadyToPackQty = 5, .PendingQty = 2, .BalanceQty = 0,
                    .Location = "B1-9", .Status = "Available", .PartNumber = "PN309", .Remarks = "Awaiting components"
                })
                availableComponents.Add(New Component With {
                    .SrNo = 10, .WID = "WID310", .CustomerName = "Cust30", .CustomerLocation = "Loc30",
                    .ProjectNo = projectName, .Description = "Cotter Pin", .QtyAvailable = 50,
                    .Color = "Silver", .ReadyToPackQty = 50, .PendingQty = 0, .BalanceQty = 0,
                    .Location = "B1-10", .Status = "Available", .PartNumber = "PN310", .Remarks = "Complete"
                })
            End If
        End Sub

        Private Sub btnCreateBundle_Click(sender As Object, e As RoutedEventArgs)
            If selectedProject Is Nothing Then
                MessageBox.Show("Please select a project first.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            If cmbBundleType.SelectedItem Is Nothing Then
                MessageBox.Show("Please select a bundle type first.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Dim bundleType As String = TryCast(cmbBundleType.SelectedItem, String)
            Dim newBundle As New Bundle With {
        .SrNo = bundles.Count + 1,
        .PackageID = $"PKG-{DateTime.Now:yyyyMMdd}-{bundles.Count + 1:000}",
        .PackageType = bundleType,
        .TotalComponents = 0,
        .Status = "New",
        .CreatedBy = Environment.UserName,
        .CreatedOn = DateTime.Now,
        .Remarks = "",
        .Components = New ObservableCollection(Of Component)()
    }

            ' Add the bundle to the list
            bundles.Add(newBundle)

            ' ✅ Auto-select the newly created bundle
            dgBundleList.SelectedItem = newBundle
            selectedBundle = newBundle
            dgBundleList.Focus()

            ' Refresh enable/disable logic
            UpdateUIState()

            MessageBox.Show($"Bundle {newBundle.PackageID} created successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        End Sub


        Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs)
            Try
                ' TODO: Add actual save logic here
                MessageBox.Show($"{bundles.Count} bundles saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Catch ex As Exception
                MessageBox.Show($"Error saving bundles: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Sub

        Private Sub btnPrintPreview_Click(sender As Object, e As RoutedEventArgs)
            If selectedBundle Is Nothing Then
                MessageBox.Show("Select a bundle first.")
                Return
            End If

            Dim firstComponent = selectedBundle.Components.FirstOrDefault()
            Dim customerName = If(firstComponent IsNot Nothing, firstComponent.CustomerName, "N/A")
            Dim customerLocation = If(firstComponent IsNot Nothing, firstComponent.CustomerLocation, "N/A")

            Dim qrData = $"Customer: {customerName}, Location: {customerLocation}, Project: {selectedProject.Name}, PackageID: {selectedBundle.PackageID}"

            Dim previewWindow As New PackingSlip()

            ' Set the window properties
            previewWindow.CustomerName = customerName
            previewWindow.CustomerLocation = customerLocation
            previewWindow.ProjectNo = selectedProject.Name
            previewWindow.PreparedDate = DateTime.Now
            previewWindow.QualityDate = DateTime.Now
            previewWindow.PackageID = selectedBundle.PackageID

            ' Find the DataGrid in the PackingSlip window and set its ItemsSource
            Dim componentGrid = TryCast(previewWindow.FindName("dgComponentList"), DataGrid)
            If componentGrid IsNot Nothing Then
                componentGrid.ItemsSource = selectedBundle.Components
            End If

            previewWindow.SetQRCode(qrData)
            previewWindow.ShowDialog()
        End Sub

        Private Sub btnAddToBundle_Click(sender As Object, e As RoutedEventArgs)
            If selectedBundle Is Nothing Then
                MessageBox.Show("Please select a bundle first.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            If dgAvailableComponents.SelectedItems.Count = 0 Then
                MessageBox.Show("Please select components to add.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Try
                Dim selectedComponents = dgAvailableComponents.SelectedItems.Cast(Of Component)().ToList()

                ' Add all selected components to bundle
                For Each comp In selectedComponents
                    availableComponents.Remove(comp)
                    comp.QtyInPackage = comp.QtyAvailable
                    selectedBundle.Components.Add(comp)
                Next

                selectedBundle.TotalComponents = selectedBundle.Components.Count
                dgBundleComponents.ItemsSource = selectedBundle.Components
                UpdateUIState()

                MessageBox.Show($"Added {selectedComponents.Count} components to bundle.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Catch ex As Exception
                MessageBox.Show($"Error adding components: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Sub

        Private Sub btnRemoveFromBundle_Click(sender As Object, e As RoutedEventArgs)
            If selectedBundle Is Nothing Then
                MessageBox.Show("Please select a bundle first.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            If dgBundleComponents.SelectedItems.Count = 0 Then
                MessageBox.Show("Please select components to remove.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            Try
                Dim selectedComponents = dgBundleComponents.SelectedItems.Cast(Of Component)().ToList()

                ' Remove all selected components from bundle
                For Each comp In selectedComponents
                    selectedBundle.Components.Remove(comp)
                    comp.QtyInPackage = 0
                    availableComponents.Add(comp)
                Next

                selectedBundle.TotalComponents = selectedBundle.Components.Count
                dgBundleComponents.ItemsSource = selectedBundle.Components
                UpdateUIState()

                MessageBox.Show($"Removed {selectedComponents.Count} components from bundle.", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            Catch ex As Exception
                MessageBox.Show($"Error removing components: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Sub

        Private Sub dgBundleList_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            selectedBundle = TryCast(dgBundleList.SelectedItem, Bundle)
            dgBundleComponents.ItemsSource = If(selectedBundle?.Components, Nothing)
            UpdateUIState()
        End Sub

        Private Sub dgAvailableComponents_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            UpdateUIState()
        End Sub

        Private Sub dgBundleComponents_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            UpdateUIState()
        End Sub
    End Class

    ' Data models
    Public Class Project
        Public Property Name As String
    End Class

    Public Class Component
        Public Property SrNo As Integer
        Public Property WID As String
        Public Property CustomerName As String
        Public Property CustomerLocation As String
        Public Property ProjectNo As String
        Public Property Description As String
        Public Property QtyAvailable As Integer
        Public Property Color As String
        Public Property ReadyToPackQty As Integer
        Public Property PendingQty As Integer
        Public Property BalanceQty As Integer
        Public Property Location As String
        Public Property Status As String
        Public Property PartNumber As String
        Public Property Remarks As String
        Public Property QtyInPackage As Integer
    End Class

    Public Class Bundle
        Public Property SrNo As Integer
        Public Property PackageID As String
        Public Property PackageType As String
        Public Property TotalComponents As Integer
        Public Property Status As String
        Public Property CreatedBy As String
        Public Property CreatedOn As DateTime
        Public Property Remarks As String
        Public Property Components As ObservableCollection(Of Component)
    End Class
End Namespace
