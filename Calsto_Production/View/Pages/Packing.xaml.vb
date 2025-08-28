Imports System.Collections.ObjectModel
Imports System.Drawing
Imports System.IO
Imports ZXing
Imports ZXing.Common


Class Packing

    Private PackingJobList As ObservableCollection(Of PackingModel)
    Private PackingBundleList As ObservableCollection(Of PackingHeaderModel)
    Private PackingBundleitemList As ObservableCollection(Of PackEntryModel)
    Private ProjectList As ObservableCollection(Of PackingProjModel)


    Public Sub New()
        InitializeComponent()
        LoadProjects()
    End Sub


    Private Sub LoadPacklist()
        Dim Packlist = PackingDBHelper.GetPacktypeList()
        CMB_Packtype.ItemsSource = Packlist
        CMB_Packtype.DisplayMemberPath = "Packtype"
        CMB_Packtype.SelectedValuePath = "Packtype"
    End Sub



    Private Sub LoadProjects()
        Dim projects = PackingDBHelper.GetProjectList()
        CMB_PackProj.ItemsSource = projects
        CMB_PackProj.DisplayMemberPath = "Proj_no"
        CMB_PackProj.SelectedValuePath = "Proj_no"
    End Sub


    Private Sub LoadpackingJobs(proNo As String)
        Try
            Dim data = PackingDBHelper.GetPackingJobs(proNo)
            PackingJobList = New ObservableCollection(Of PackingModel)(data)
            ProjDG.ItemsSource = PackingJobList
        Catch ex As Exception
            MessageBox.Show("Error loading project list: " & ex.Message)
        End Try
    End Sub


    Private Sub LoadBundles(proNo As String)
        Try
            Dim data = PackingDBHelper.GetBundleList(proNo)
            PackingBundleList = New ObservableCollection(Of PackingHeaderModel)(data)
            BundleDG.ItemsSource = PackingBundleList
        Catch ex As Exception
            MessageBox.Show("Error loading project list: " & ex.Message)
        End Try
    End Sub


    Private Sub Applybtn_Click(sender As Object, e As RoutedEventArgs)


        Dim selectedProj = TryCast(CMB_PackProj.SelectedItem, PackingProjModel)
        If selectedProj IsNot Nothing Then
            Txt_customer.Text = selectedProj.Customer
            LoadpackingJobs(selectedProj.Proj_no)
            LoadBundles(selectedProj.Proj_no)
        End If
        ProjDG.IsEnabled = True
        Btn_Projapply.IsEnabled = False
        PROJPANEL.IsEnabled = True
        CMB_PackProj.IsEnabled = False
        Btn_ProjClear.IsEnabled = True
        LoadPacklist()
    End Sub


    Private Sub Clearbtn_Click(sender As Object, e As RoutedEventArgs)

        Txt_customer.Text = ""
        CMB_Packtype.SelectedItem = ""
        ProjDG.ItemsSource = Nothing
        BundleDG.ItemsSource = Nothing
        BundleitemDG.ItemsSource = Nothing
        ProjDG.IsEnabled = False
        Btn_Projapply.IsEnabled = True
        PROJPANEL.IsEnabled = False
        CMB_PackProj.IsEnabled = True
        Btn_ProjClear.IsEnabled = False
        Editexit()
    End Sub


    Private Sub BundleDG_sel_change()

        Dim selectedBundle As PackingHeaderModel = TryCast(BundleDG.SelectedItem, PackingHeaderModel)


        If selectedBundle IsNot Nothing Then


            Try
                Dim data = PackingDBHelper.GetPackEntryList(selectedBundle.PackID)
                PackingBundleitemList = New ObservableCollection(Of PackEntryModel)(data)
                BundleitemDG.ItemsSource = PackingBundleitemList
            Catch ex As Exception
                MessageBox.Show("Error loading project list: " & ex.Message)
            End Try
        Else


        End If

    End Sub

    Private Sub Bundledit_Click(sender As Object, e As RoutedEventArgs)

        Dim selectedBundle As PackingHeaderModel = TryCast(BundleDG.SelectedItem, PackingHeaderModel)


        If selectedBundle IsNot Nothing Then

            Try
                Dim data = PackingDBHelper.GetPackEntryList(selectedBundle.PackID)
                PackingBundleitemList = New ObservableCollection(Of PackEntryModel)(data)
                BundleitemDG.ItemsSource = PackingBundleitemList
            Catch ex As Exception
                MessageBox.Show("Error loading project list: " & ex.Message)
            End Try
            Bundleeditpanel.IsEnabled = True
            BundleDG.IsEnabled = False
            exitpack.IsEnabled = True
            createbundlepanel.isEnabled = False
        Else


        End If

    End Sub



    Private Sub Projitemfocus()
        bndl_apply.IsEnabled = True
        bndl_qty.IsEnabled = True
        bndl_remove.IsEnabled = False
    End Sub


    Private Sub BundleitemFocus()
        bndl_apply.IsEnabled = False
        bndl_qty.IsEnabled = False
        bndl_qty.Text = ""
        bndl_remove.IsEnabled = True

    End Sub



    Private Sub Editexit()
        Bundleeditpanel.IsEnabled = False
        bndl_qty.Text = ""
        BundleDG.IsEnabled = True
        exitpack.IsEnabled = False
        createbundlepanel.IsEnabled = True
        BundleitemDG.ItemsSource = Nothing

    End Sub




    Private Sub CreateBundle_Click(sender As Object, e As RoutedEventArgs)

        Dim selectedProj = TryCast(CMB_PackProj.SelectedItem, PackingProjModel)
        Dim selectedPacktype = TryCast(CMB_Packtype.SelectedItem, PackingTypeModel)
        Dim createdBy As String = Environment.UserName
        Dim Remarks As String = ""


        If selectedPacktype Is Nothing Then
            MessageBox.Show("Please select a Pack Type before creating a bundle.",
                            "Missing Pack Type", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Save to DB
        PackingDBHelper.CreateBundle(selectedPacktype.Packtype, selectedProj.Proj_no, createdBy)

        ' Success message
        MessageBox.Show($"Lot entry applied for Project {selectedProj.Proj_no} with Pack Type {selectedPacktype.Packtype} successfully.",
                        "Lot Entry", MessageBoxButton.OK, MessageBoxImage.Information)



        If selectedProj IsNot Nothing Then
            Txt_customer.Text = selectedProj.Customer

            LoadBundles(selectedProj.Proj_no)
        End If

    End Sub




    Private Sub Additem_Click(sender As Object, e As RoutedEventArgs)


        Dim createdBy As String = Environment.UserName
        Dim Remarks As String = ""
        Dim SelectedProject = TryCast(ProjDG.SelectedItem, PackingModel)
        Dim wid = SelectedProject.WID
        Dim Description = SelectedProject.Description
        Dim JCno = SelectedProject.JC_no
        Dim qTY = bndl_qty.Text
        Dim selectedBundle As PackingHeaderModel = TryCast(BundleDG.SelectedItem, PackingHeaderModel)
        Dim Pid = selectedBundle.PackID
        Dim Projno = selectedBundle.ProjNo
        If bndl_qty.Text Is Nothing Then
            MessageBox.Show("Please give valid Qty before Adding a Item.",
                            "Valid Qty needed.", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If


        ' Save to DB
        PackingDBHelper.Additemstobundle(wid, qTY, Pid, Projno, JCno, createdBy)

        ' Success message
        MessageBox.Show($"Item:{Description} added to Bundle :{Pid} successfully.",
                        "Add items to bundle", MessageBoxButton.OK, MessageBoxImage.Information)
        BundleDG_sel_change()
        RefreshBOM()

    End Sub


    Private Sub Removeitem_Click(sender As Object, e As RoutedEventArgs)

        ' Validate Bundle selection
        Dim selectedBundle As PackingHeaderModel = TryCast(BundleDG.SelectedItem, PackingHeaderModel)
        If selectedBundle Is Nothing Then
            MessageBox.Show("Please select a bundle before removing an item.",
                            "Bundle not selected", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Validate Project/Entry selection
        Dim SelectedProject As PackEntryModel = TryCast(BundleitemDG.SelectedItem, PackEntryModel)
        If SelectedProject Is Nothing Then
            MessageBox.Show("Please select a project/item before removing it from the bundle.",
                            "Item not selected", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim Pid = selectedBundle.PackID
        Dim Description = SelectedProject.Description
        Dim PackEntryID = SelectedProject.ID

        PackingDBHelper.Removeitems(PackEntryID)

        ' Success message
        MessageBox.Show($"Item: {Description} removed from Bundle: {Pid} successfully.",
                        "Remove Item", MessageBoxButton.OK, MessageBoxImage.Information)
        BundleDG_sel_change()
        RefreshBOM()

    End Sub



    Private Sub RefreshBOM()
        Try
            ' 1. Save current grid preset
            ProjDG.SavePreset()

            ' 2. Remember currently selected WID
            Dim selectedWID As String = Nothing
            If ProjDG.SelectedItem IsNot Nothing Then
                selectedWID = CType(ProjDG.SelectedItem, PackingModel).WID
            End If

            ' 3. Get current project number
            Dim selectedProj = TryCast(CMB_PackProj.SelectedItem, PackingProjModel)
            If selectedProj Is Nothing Then Exit Sub
            Dim proNo = selectedProj.Proj_no

            ' 4. Reload BOM items from database
            Dim data = PackingDBHelper.GetPackingJobs(proNo)
            PackingJobList = New ObservableCollection(Of PackingModel)(data)
            ProjDG.ItemsSource = PackingJobList

            ' 5. Restore grid preset
            ProjDG.LoadPreset()

            ' 6. Determine which item to select
            Dim rowToSelect = PackingJobList.FirstOrDefault(Function(x) x.WID = selectedWID)

            If rowToSelect IsNot Nothing Then
                ' 7. Select row and scroll into view
                ProjDG.SelectedItem = rowToSelect
                ProjDG.UpdateLayout()
                ProjDG.ScrollIntoView(rowToSelect)


            End If

        Catch ex As Exception
            MessageBox.Show("Error refreshing BOM: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub




    Private Sub openpackingslip()

        Dim selectedBundle As PackingHeaderModel = TryCast(BundleDG.SelectedItem, PackingHeaderModel)
        Dim packID = selectedBundle.PackID

        Dim header = PackingSlipDBHelper.GetPackingHeader(packID)
        If header IsNot Nothing Then
            header.Items = PackingSlipDBHelper.GetPackingItems(packID)

            ' Pass to your FlowDocument Window
            Dim slipWin As New PackingSlipWindow(header)
            slipWin.ShowDialog()
        Else
            MessageBox.Show("No data found for PackID " & packID)
        End If

    End Sub




End Class

