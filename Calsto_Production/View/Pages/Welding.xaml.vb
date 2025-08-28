Imports System.Collections.ObjectModel

Public Class Welding
    Inherits Page

    Private WeldingJobList As ObservableCollection(Of WeldingModel)
    Private LotList As ObservableCollection(Of WeldingLotModel)
    Private ProjectList As ObservableCollection(Of WeldingProjModel)

    Public Sub New()
        InitializeComponent()
        LoadProjects()
    End Sub


    Private Sub LoadProjects()
        Try
            Dim data = WeldingDBHelper.GetProjectNO()
            ProjectList = New ObservableCollection(Of WeldingProjModel)(data)
            ProjDG.ItemsSource = ProjectList
        Catch ex As Exception
            MessageBox.Show("Error loading project list: " & ex.Message)
        End Try
    End Sub

    ' Load all Welding jobs
    Private Sub LoadWeldingJobs(proNo As String)
        Try
            Dim data = WeldingDBHelper.GetWeldingJobs(proNo)
            WeldingJobList = New ObservableCollection(Of WeldingModel)(data)
            FormDG.ItemsSource = WeldingJobList
        Catch ex As Exception
            MessageBox.Show("Error loading project list: " & ex.Message)
        End Try
    End Sub


    Private Sub LoadLotIDs(jobc As String)
        Try
            Dim data = WeldingDBHelper.GetWeldingLots(jobc)
            LotList = New ObservableCollection(Of WeldingLotModel)(data)
            JobcardDG.ItemsSource = LotList
        Catch ex As Exception
            MessageBox.Show("Error loading project side details: " & ex.Message)
        End Try
    End Sub




    Private Sub FormProj_Sel_Changed(sender As Object, e As SelectionChangedEventArgs)
        Dim selectedProj As WeldingProjModel = TryCast(ProjDG.SelectedItem, WeldingProjModel)
        If selectedProj IsNot Nothing Then
            LoadWeldingJobs(selectedProj.PROJECTNO)
            txt_Projectno.Text = selectedProj.PROJECTNO
            txt_Customer.Text = selectedProj.Customer
            Txt_ReqQty.Text = ""
            Txt_BalQty.Text = ""
            Txt_Colour.Text = ""
            Txt_Description.Text = ""
            Txt_CompQty.Text = ""
            DetailsTab.IsEnabled = False
            JobcardDG.ItemsSource = Nothing

        End If
    End Sub


    Private Sub dgLOT_Sel_Changed(sender As Object, e As SelectionChangedEventArgs)
        Dim selectedlot As WeldingLotModel = TryCast(JobcardDG.SelectedItem, WeldingLotModel)
        If selectedlot IsNot Nothing Then
            print_btn.IsEnabled = True
            editbtn.IsEnabled = True
            editbtn.Visibility = Visibility.Visible
            Lotedit.Text = ""
            Lotedit.Visibility = Visibility.Collapsed
            Lotapply.Visibility = Visibility.Collapsed
            editbtn.Visibility = Visibility.Visible

        Else
            print_btn.IsEnabled = False
            editbtn.IsEnabled = False
            editbtn.Visibility = Visibility.Visible
            editcancel.Visibility = Visibility.Collapsed

        End If

    End Sub


    Private Sub Editbtn_Click(sender As Object, e As RoutedEventArgs)
        Dim selectedlot As WeldingLotModel = TryCast(JobcardDG.SelectedItem, WeldingLotModel)

        If selectedlot IsNot Nothing Then
            Lotedit.IsEnabled = True
            Lotedit.Visibility = Visibility.Visible
            Lotedit.Text = selectedlot.Quantity
            Lotapply.IsEnabled = True
            Lotapply.Visibility = Visibility.Visible
            editbtn.Visibility = Visibility.Collapsed
            editcancel.IsEnabled = True
            editcancel.Visibility = Visibility.Visible

        End If

    End Sub



    Private Sub Cancel_Click(sender As Object, e As RoutedEventArgs)
        Lotedit.Visibility = Visibility.Collapsed
        Lotapply.Visibility = Visibility.Collapsed
        editbtn.Visibility = Visibility.Visible

    End Sub


    Private Sub Change_lot_click(sender As Object, e As RoutedEventArgs)

        If JobcardDG.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a job to Update.")
            Exit Sub
        End If
        Dim selectedlot As WeldingLotModel = TryCast(JobcardDG.SelectedItem, WeldingLotModel)

        If MessageBox.Show($"Update the Qtny for Lot {selectedlot.JC_no}?",
                       "Confirm Update",
                       MessageBoxButton.YesNo,
                       MessageBoxImage.Warning) = MessageBoxResult.No Then
            Exit Sub
        End If

        Dim transid As String = selectedlot.Transaction_ID
        Dim Lotid As String = selectedlot.JC_no
        Dim newqty As String = Lotedit.Text
        Dim createdBy As String = Environment.UserName
        Dim Remarks As String = "Update on:" & Now.ToString("yyyy-MM-dd HH:mm:ss")

        WeldingDBHelper.EditLotEntry(transid, Lotid, newqty, createdBy, Remarks)

        Dim selectedJC As WeldingModel = TryCast(FormDG.SelectedItem, WeldingModel)
        If selectedJC IsNot Nothing Then
            LoadLotIDs(selectedJC.JC_no)
            Txt_ReqQty.Text = selectedJC.BOMQty
            Txt_BalQty.Text = selectedJC.BalanceQty
            Txt_Colour.Text = selectedJC.Colour
            Txt_Description.Text = selectedJC.Description
            Txt_CompQty.Text = selectedJC.ProducedQty

            If selectedJC.Status = "CLOSED" Then
                DetailsTab.IsEnabled = False
            Else
                DetailsTab.IsEnabled = True
            End If
        End If

        Lotedit.Visibility = Visibility.Collapsed
        Lotapply.Visibility = Visibility.Collapsed
        editbtn.Visibility = Visibility.Visible
        Lotedit.Text = ""

        RefreshBOM()
        Refreshproj()

        Lot_qty.Text = ""

    End Sub






    Private Sub dgFORM_Sel_Changed(sender As Object, e As SelectionChangedEventArgs)
        Dim selectedJC As WeldingModel = TryCast(FormDG.SelectedItem, WeldingModel)
        If selectedJC IsNot Nothing Then
            LoadLotIDs(selectedJC.JC_no)
            Txt_ReqQty.Text = selectedJC.BOMQty
            Txt_BalQty.Text = selectedJC.BalanceQty
            Txt_Colour.Text = selectedJC.Colour
            Txt_Description.Text = selectedJC.Description
            Txt_CompQty.Text = selectedJC.ProducedQty

            If selectedJC.Status = "CLOSED" Then
                DetailsTab.IsEnabled = False
            Else
                DetailsTab.IsEnabled = True
            End If
        End If

        Lotedit.Visibility = Visibility.Collapsed
        Lotapply.Visibility = Visibility.Collapsed
        editbtn.Visibility = Visibility.Visible
        Lotedit.Text = ""


    End Sub


    Private Sub LotEntry_Click(sender As Object, e As RoutedEventArgs)
        If FormDG.SelectedItem Is Nothing Then
            MessageBox.Show("Please select a job to apply the lot entry.", "Lot Entry", MessageBoxButton.OK, MessageBoxImage.Warning)
            Exit Sub
        End If

        Dim selectedJC As WeldingModel = CType(FormDG.SelectedItem, WeldingModel)

        ' Confirm
        If MessageBox.Show($"Apply the Qty for this item: {selectedJC.JC_no}?",
                           "Confirm Lot Entry",
                           MessageBoxButton.YesNo,
                           MessageBoxImage.Question) = MessageBoxResult.No Then
            Exit Sub
        End If

        ' Validate Qty
        Dim JCQty As Integer
        If Not Integer.TryParse(Lot_qty.Text, JCQty) OrElse JCQty <= 0 Then
            MessageBox.Show("Please enter a valid quantity.", "Invalid Entry", MessageBoxButton.OK, MessageBoxImage.Error)
            Exit Sub
        End If

        ' 🔹 Validate against Balance Qty
        Dim balanceQty As Integer
        If Not Integer.TryParse(Txt_BalQty.Text, balanceQty) Then
            balanceQty = 0
        End If

        If JCQty > balanceQty Then
            MessageBox.Show($"Lot entry quantity ({JCQty}) cannot be greater than Balance Qty ({balanceQty}).",
                            "Lot Entry Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error)
            Exit Sub
        End If

        Dim createdBy As String = Environment.UserName
        Dim Remarks As String = ""

        ' Save to DB
        WeldingDBHelper.CreateJobEntry(selectedJC.JC_no, JCQty, createdBy, Remarks)
        RefreshBOM()
        Refreshproj()
        Lot_qty.Text = ""

        MessageBox.Show($"Lot entry applied for Job Card {selectedJC} successfully.", "Lot Entry", MessageBoxButton.OK, MessageBoxImage.Information)
    End Sub


    Private Sub RefreshBOM()
        Try
            ' 1. Save current grid preset
            FormDG.SavePreset()

            ' 2. Remember currently selected WID
            Dim selectedWID As String = Nothing
            If FormDG.SelectedItem IsNot Nothing Then
                selectedWID = CType(FormDG.SelectedItem, WeldingModel).WID
            End If

            ' 3. Get current project number
            Dim selectedProj = CType(ProjDG.SelectedItem, WeldingProjModel)
            If selectedProj Is Nothing Then Exit Sub
            Dim proNo = selectedProj.PROJECTNO

            ' 4. Reload BOM items from database
            Dim data = WeldingDBHelper.GetWeldingJobs(proNo)
            WeldingJobList = New ObservableCollection(Of WeldingModel)(data)
            FormDG.ItemsSource = WeldingJobList

            ' 5. Restore grid preset
            FormDG.LoadPreset()

            ' 6. Determine which item to select
            Dim rowToSelect = WeldingJobList.FirstOrDefault(Function(x) x.WID = selectedWID)

            If rowToSelect IsNot Nothing Then
                ' 7. Select row and scroll into view
                FormDG.SelectedItem = rowToSelect
                FormDG.UpdateLayout()
                FormDG.ScrollIntoView(rowToSelect)

                ' 8. Update detail fields safely
                Txt_ReqQty.Text = If(rowToSelect.BOMQty.ToString(), "0")
                Txt_BalQty.Text = If(rowToSelect.BalanceQty.ToString(), "0")
                Txt_Colour.Text = If(rowToSelect.Colour, String.Empty)
                Txt_Description.Text = If(rowToSelect.Description, String.Empty)
                Txt_CompQty.Text = If(rowToSelect.ProducedQty.ToString(), "0")

                ' 9. Enable or disable Details tab based on status
                DetailsTab.IsEnabled = Not String.Equals(rowToSelect.Status, "CLOSED", StringComparison.OrdinalIgnoreCase)
            End If

        Catch ex As Exception
            MessageBox.Show("Error refreshing BOM: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub






    Private Sub Refreshproj()
        Try
            ' ==== Refresh Projects ====
            ProjDG.SavePreset()

            Dim selectedProjNo As String = Nothing
            If ProjDG.SelectedItem IsNot Nothing Then
                selectedProjNo = CType(ProjDG.SelectedItem, WeldingProjModel).PROJECTNO
            End If

            Dim projData = WeldingDBHelper.GetProjectNO()
            ProjectList = New ObservableCollection(Of WeldingProjModel)(projData)
            ProjDG.ItemsSource = ProjectList

            ProjDG.LoadPreset()

            ' Restore project selection
            Dim currentProjNo As String = Nothing
            If selectedProjNo IsNot Nothing Then
                Dim projRow = ProjectList.FirstOrDefault(Function(x) x.PROJECTNO = selectedProjNo)
                If projRow IsNot Nothing Then
                    ProjDG.SelectedItem = projRow
                    ProjDG.UpdateLayout()
                    ProjDG.ScrollIntoView(projRow)
                    currentProjNo = projRow.PROJECTNO

                End If


            End If

        Catch ex As Exception
            MessageBox.Show("Error refreshing: " & ex.Message)
        End Try
    End Sub





End Class
